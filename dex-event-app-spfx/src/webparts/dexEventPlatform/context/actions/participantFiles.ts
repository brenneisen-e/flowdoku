/**
 * Teilnehmer-Register, Event-Nummern und Datei-Anhaenge (eigene Uploads und
 * Feld-Dokumente je Anmeldung).
 *
 * v30.66: Aus `EventContext.tsx` herausgezogen (Modularisierung Stufe 3).
 * Die Funktionskoerper sind unveraendert - statt aus der Provider-Closure
 * beziehen sie ihre Umgebung aus dem `deps`-Objekt.
 */

import { EventService, SPParticipant } from '../../services/EventService';

export interface ParticipantFileDeps {
  eventService: EventService;
  subsiteMap: { current: Record<string, string> };
  currentUserEmail: string;
  loadEvents: () => Promise<void>;
}

export function makeParticipantFileActions(deps: ParticipantFileDeps) {
  const { eventService, subsiteMap, currentUserEmail, loadEvents } = deps;

  // v22.50: Passthrough für die globale Admin-Suche (DEX_Participants).
  async function getAllParticipants(): Promise<SPParticipant[]> {
    try { return await eventService.getAllParticipants(); }
    catch (err) { console.warn('[DEX] getAllParticipants failed:', err); return []; }
  }

  // v28.22: Event-Nummern (angemeldet / Warteliste) einer BELIEBIGEN Person aus
  // DEX_Participants. Diese Liste liegt auf der Haupt-Site und unterliegt NICHT
  // der Item-Level-Security der Teilnehmerlisten — sie kennt daher auch
  // Anmeldungen, die jemand anders (z.B. eine Assistenz) angelegt hat und die
  // für die betroffene Person selbst unsichtbar sind. Genau daraus entstanden
  // die doppelten Anmeldungen: Der Vorab-Check auf der Event-Subsite fand die
  // fremde Zeile nicht und legte fail-open eine zweite an.
  async function getEventNumbersForEmail(email: string): Promise<{ registered: number[]; waitlisted: number[] }> {
    const em = (email || '').trim();
    if (!em) return { registered: [], waitlisted: [] };
    try {
      const record = await eventService.getParticipantByEmail(em);
      if (!record) return { registered: [], waitlisted: [] };
      const parse = (s?: string): number[] => s
        ? s.split(',').map(x => parseInt(x.trim(), 10)).filter(n => !isNaN(n))
        : [];
      return { registered: parse(record.EventRegistered), waitlisted: parse(record.EventOnWaitlist) };
    } catch {
      return { registered: [], waitlisted: [] };
    }
  }

  async function getMyEventNumbers(): Promise<{ registered: number[]; waitlisted: number[] }> {
    try {
      const record = await eventService.getParticipantByEmail(currentUserEmail);
      if (!record) return { registered: [], waitlisted: [] };
      const registered = record.EventRegistered
        ? record.EventRegistered.split(',').map(s => parseInt(s.trim(), 10)).filter(n => !isNaN(n))
        : [];
      const waitlisted = record.EventOnWaitlist
        ? record.EventOnWaitlist.split(',').map(s => parseInt(s.trim(), 10)).filter(n => !isNaN(n))
        : [];
      return { registered, waitlisted };
    } catch {
      return { registered: [], waitlisted: [] };
    }
  }

  async function refreshEvents(): Promise<void> {
    await loadEvents();
  }

  // v11.0: Item-Attachments — Wrapper für die eigene Registrierung.
  async function listMyEventAttachments(eventId: string): Promise<Array<{ fileName: string; serverRelativeUrl: string }>> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return [];
    const myReg = await eventService.getMyRegistration(subsiteUrl, currentUserEmail);
    if (!myReg) return [];
    return eventService.listRegistrationAttachments(subsiteUrl, myReg.Id);
  }
  async function uploadMyEventAttachment(eventId: string, file: File): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;
    const myReg = await eventService.getMyRegistration(subsiteUrl, currentUserEmail);
    if (!myReg) return false;
    return eventService.addRegistrationAttachment(subsiteUrl, myReg.Id, file);
  }
  async function deleteMyEventAttachment(eventId: string, fileName: string): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;
    const myReg = await eventService.getMyRegistration(subsiteUrl, currentUserEmail);
    if (!myReg) return false;
    return eventService.deleteRegistrationAttachment(subsiteUrl, myReg.Id, fileName);
  }

  // v19.0: Dokument-Custom-Felder. Ein Attachment wird über einen Dateinamen-
  // Präfix (`dxf-<fieldId>--`) genau EINEM Dokument-Feld zugeordnet, sodass ein
  // Event mehrere Dokument-Felder haben kann. participantEmail erlaubt den
  // Upload für eine andere Person (stellvertretende Anmeldung); Default = der
  // eingeloggte User (Self-Anmeldung + „Meine Events").
  const docFieldPrefix = (fieldId: string): string => `dxf-${(fieldId || '').replace(/[^a-zA-Z0-9]/g, '')}--`;
  const stripDocPrefix = (fileName: string): string =>
    fileName
      .replace(/^dxf-[a-zA-Z0-9]+--\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}_/, '')
      .replace(/^dxf-[a-zA-Z0-9]+--/, '');
  async function uploadFieldDocument(eventId: string, fieldId: string, file: File, participantEmail?: string): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;
    const email = (participantEmail || currentUserEmail || '').trim();
    const reg = await eventService.getMyRegistration(subsiteUrl, email);
    if (!reg || !reg.Id) return false;
    return eventService.addRegistrationAttachment(subsiteUrl, reg.Id, file, docFieldPrefix(fieldId));
  }
  async function listFieldDocuments(eventId: string, fieldId: string, participantEmail?: string): Promise<Array<{ fileName: string; serverRelativeUrl: string; displayName: string }>> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return [];
    const email = (participantEmail || currentUserEmail || '').trim();
    const reg = await eventService.getMyRegistration(subsiteUrl, email);
    if (!reg || !reg.Id) return [];
    const all = await eventService.listRegistrationAttachments(subsiteUrl, reg.Id);
    const prefix = docFieldPrefix(fieldId);
    return all.filter(f => f.fileName.startsWith(prefix)).map(f => ({ ...f, displayName: stripDocPrefix(f.fileName) }));
  }
  async function deleteFieldDocument(eventId: string, fileName: string, participantEmail?: string): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;
    const email = (participantEmail || currentUserEmail || '').trim();
    const reg = await eventService.getMyRegistration(subsiteUrl, email);
    if (!reg || !reg.Id) return false;
    return eventService.deleteRegistrationAttachment(subsiteUrl, reg.Id, fileName);
  }

  return { getAllParticipants, getEventNumbersForEmail, getMyEventNumbers, refreshEvents, listMyEventAttachments, uploadMyEventAttachment, deleteMyEventAttachment, uploadFieldDocument, listFieldDocuments, deleteFieldDocument };
}
