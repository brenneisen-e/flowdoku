/* useTeamActions — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 2469-2619 des
 * Stands vor dem Schnitt). Der Rumpf ist zeichengleich uebernommen; was die
 * Gruppe aus dem Komponenten-Scope liest, kommt als `ctx` herein, was sie
 * nach aussen liefert, geht als Objekt zurueck.
 */
import * as React from 'react';
import { EventService, SPRegistration } from '../../../services/EventService';
import { wrapTemplate } from '../../../services/EmailTemplates';
import { DeloitteEvent } from '../../../types';

export interface UseTeamActionsCtx {
  dragRegId: number;
  eventServiceRef: EventService;
  /** v30.67 (Review): gemeinsamer Nachlade-Pfad der Seite — `null` = nicht lesbar. */
  reloadRegistrations: () => Promise<SPRegistration[] | null>;
  idFixCheckedForRef: React.MutableRefObject<string>;
  isDe: boolean;
  isLoadingRegs: boolean;
  recentCancellation: (regs: SPRegistration[]) => {    recent: boolean;    whenIso: string;    detail: string;};
  registrations: SPRegistration[];
  reloadRegistrationsForIdCheck: () => Promise<void>;
  selectedEvent: DeloitteEvent;
  setDragOverTid: React.Dispatch<React.SetStateAction<string>>;
  setDragRegId: React.Dispatch<React.SetStateAction<number>>;
  setTeamMailBody: React.Dispatch<React.SetStateAction<string>>;
  setTeamMailInfoByTid: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  setTeamMailOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamMailSending: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamMailSubject: React.Dispatch<React.SetStateAction<string>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  teamMailBody: string;
  teamMailInfoByTid: Record<string, string>;
  teamMailSubject: string;
}

export interface UseTeamActionsResult {
  columnOrder: string[];
  getActiveTeams: () => Array<{    tid: string;    teamName: string;    members: SPRegistration[];}>;
  hiddenColumns: string[];
  moveRegToTeam: (reg: SPRegistration, targetTid: string, targetTeamName: string | undefined) => Promise<void>;
  onTeamDrop: (targetTid: string, targetTeamName: string | undefined) => void;
  openTeamMailDialog: () => void;
  sendTeamMails: () => Promise<void>;
  setColumnOrder: React.Dispatch<React.SetStateAction<string[]>>;
  setHiddenColumns: React.Dispatch<React.SetStateAction<string[]>>;
  setShowColumnPicker: React.Dispatch<React.SetStateAction<boolean>>;
  setShowMatches: React.Dispatch<React.SetStateAction<boolean>>;
  showColumnPicker: boolean;
  showMatches: boolean;
}

export function useTeamActions(ctx: UseTeamActionsCtx): UseTeamActionsResult {
  const {
    dragRegId, eventServiceRef, idFixCheckedForRef, isDe, isLoadingRegs,
    recentCancellation, registrations, reloadRegistrations, reloadRegistrationsForIdCheck, selectedEvent,
    setDragOverTid, setDragRegId, setTeamMailBody, setTeamMailInfoByTid,
    setTeamMailOpen, setTeamMailSending, setTeamMailSubject, showAlert, teamMailBody,
    teamMailInfoByTid, teamMailSubject,
  } = ctx;
  // v23.0: Eine Registrierung per Drag&Drop in ein Team/eine Break-Out-Session
  // (targetTid) oder zurück „ohne Team" ('') verschieben. War die Person Lead
  // ihres alten Teams und bleiben Mitglieder, rückt die früheste nach.
  const moveRegToTeam = async (reg: SPRegistration, targetTid: string, targetTeamName: string | undefined): Promise<void> => {
    if (!selectedEvent?.subsiteUrl || !eventServiceRef) return;
    const sub = selectedEvent.subsiteUrl;
    const curTid = reg.TeamId || '';
    if (curTid === (targetTid || '')) return;
    try {
      await eventServiceRef.assignRegistrationToTeam(sub, reg.Id, targetTid || '', targetTeamName || '', false);
      if (curTid && reg.TeamLead) {
        const rest = registrations.filter(x => x.Id !== reg.Id && (x.TeamId || '') === curTid && x.Status !== 'Abgemeldet');
        if (rest.length > 0) {
          rest.sort((a, b) => ((a.TeilnehmerID ?? 9_999_999) as number) - ((b.TeilnehmerID ?? 9_999_999) as number));
          const tn = rest.find(x => x.TeamName)?.TeamName || '';
          try { await eventServiceRef.assignRegistrationToTeam(sub, rest[0].Id, curTid, tn || '', true); } catch { /* */ }
        }
      }
      const nm = `${reg.Vorname || ''} ${reg.Nachname || ''}`.trim() || reg.ParticipantName || reg.ParticipantEmail;
      eventServiceRef.writeChangeLog({
        action: targetTid ? 'TeamMemberAssigned' : 'TeamMemberRemoved',
        targetType: 'Participant', targetId: reg.ParticipantEmail, targetName: nm,
        eventId: selectedEvent.id, eventTitle: selectedEvent.title,
        details: { fromTeam: curTid, toTeam: targetTid, via: 'dragdrop' },
      }).catch(() => { /* */ });
      // v30.67 (Review): gemeinsamer Nachlade-Pfad statt `[]` bei 429.
      await reloadRegistrations();
    } catch (err) { console.warn('[DEX] moveRegToTeam failed:', err); }
  };
  // Drop-Handler: gezogene Registrierung ermitteln + verschieben.
  const onTeamDrop = (targetTid: string, targetTeamName: string | undefined): void => {
    setDragOverTid(null);
    const id = dragRegId;
    setDragRegId(null);
    if (id === null) return;
    const reg = registrations.find(r => r.Id === id);
    if (reg) moveRegToTeam(reg, targetTid, targetTeamName).catch(() => { /* */ });
  };

  // v23.0: Aktive Teams aus den geladenen Registrierungen gruppieren
  // (für die Per-Team-Mail).
  const getActiveTeams = (): Array<{ tid: string; teamName: string; members: SPRegistration[] }> => {
    const map: Record<string, SPRegistration[]> = {};
    for (const r of registrations) {
      if (r.Status === 'Abgemeldet') continue;
      const tid = r.TeamId || '';
      if (!tid) continue;
      (map[tid] = map[tid] || []).push(r);
    }
    return Object.entries(map).map(([tid, members]) => ({ tid, members, teamName: members.find(m => m.TeamName)?.TeamName || '' }));
  };
  // Mail-Dialog mit vorausgefülltem Text öffnen.
  const openTeamMailDialog = (): void => {
    if (!selectedEvent) return;
    const termS = selectedEvent.teamTermSingular || 'Team';
    setTeamMailSubject(isDe ? `Deine ${termS}: ${selectedEvent.title}` : `Your ${termS}: ${selectedEvent.title}`);
    setTeamMailBody(isDe
      ? `<p>Hallo {{Vorname}},</p>\n<p>hier sind die Infos zu deiner <strong>${termS}</strong> beim Event <strong>{{EventTitle}}</strong>:</p>\n<p><strong>{{TeamName}}</strong></p>\n<p>{{TeamInfo}}</p>\n<p>Viele Grüße<br />Dein Event-Team</p>`
      : `<p>Hi {{Vorname}},</p>\n<p>here is the info for your <strong>${termS}</strong> at <strong>{{EventTitle}}</strong>:</p>\n<p><strong>{{TeamName}}</strong></p>\n<p>{{TeamInfo}}</p>\n<p>Best regards<br />Your event team</p>`);
    const init: Record<string, string> = {};
    for (const t of getActiveTeams()) init[t.tid] = teamMailInfoByTid[t.tid] || '';
    setTeamMailInfoByTid(init);
    setTeamMailOpen(true);
  };
  // Pro Team: jedes aktive Mitglied bekommt eine eigene Mail mit team-
  // spezifischer Info (z.B. Teams-Einwahllink). Im Deloitte-Layout gewrappt.
  const sendTeamMails = async (): Promise<void> => {
    if (!selectedEvent || !eventServiceRef) return;
    if (selectedEvent.disableEmails) {
      showAlert(isDe ? 'E-Mails sind für dieses Event deaktiviert (Schritt 6 „Kommunikation").' : 'Emails are disabled for this event (step 6 “Communication”).', { variant: 'error' });
      return;
    }
    setTeamMailSending(true);
    const escHtml = (s: string): string => (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    // URLs in der Team-Info klickbar machen + Zeilenumbrüche zu <br>.
    const linkify = (raw: string): string => escHtml(raw)
      .replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" style="color:#86bc25;font-weight:600;">$1</a>')
      .replace(/\n/g, '<br />');
    const termS = selectedEvent.teamTermSingular || 'Team';
    let sent = 0;
    for (const t of getActiveTeams()) {
      const infoHtml = linkify((teamMailInfoByTid[t.tid] || '').trim());
      const tName = t.teamName || termS;
      for (const m of t.members) {
        const first = (m.Vorname && m.Vorname.trim()) || (m.ParticipantName || '').split(/\s+/)[0] || '';
        const fullName = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantName || m.ParticipantEmail;
        const bodyFilled = teamMailBody
          .replace(/\{\{Vorname\}\}/g, escHtml(first))
          .replace(/\{\{Name\}\}/g, escHtml(fullName))
          .replace(/\{\{TeamName\}\}/g, escHtml(tName))
          .replace(/\{\{EventTitle\}\}/g, escHtml(selectedEvent.title))
          .replace(/\{\{TeamInfo\}\}/g, infoHtml || (isDe ? '<em>(keine zusätzlichen Infos)</em>' : '<em>(no additional info)</em>'));
        const subjectFilled = teamMailSubject
          .replace(/\{\{TeamName\}\}/g, tName)
          .replace(/\{\{EventTitle\}\}/g, selectedEvent.title);
        const wrapped = wrapTemplate('#86bc25', subjectFilled, tName, bodyFilled);
        try {
          const ok = await eventServiceRef.queueEmail(subjectFilled, m.ParticipantEmail, fullName, wrapped, 'TeamInfo', selectedEvent.title, selectedEvent.id);
          if (ok) sent += 1;
        } catch { /* best-effort pro Empfänger */ }
      }
    }
    setTeamMailSending(false);
    setTeamMailOpen(false);
    showAlert(isDe ? `${sent} Mail(s) in die Warteschlange gelegt — sie werden in Kürze versendet.` : `${sent} mail(s) queued — they will be sent shortly.`, { variant: 'success' });
  };

  // Max. 10 automatische Neu-Checks (≈5 Min) pro Event — wenn die Lücke dann
  // immer noch da ist, ist sie echt (Tail-Race, siehe Box-Text) und kein
  // weiteres Polling nötig.
  const idRecheckCountRef = React.useRef(0);
  React.useEffect(() => { idRecheckCountRef.current = 0; }, [selectedEvent?.id]);
  React.useEffect(() => {
    if (!selectedEvent) return undefined;
    // v22.67: kein ID-Durchgängigkeits-Polling im Klammer-Modus (Schatten-Zeilen
    // haben keine fortlaufenden Nummern — das war ein Fehlalarm).
    if (selectedEvent.subEventsOnlyMode) return undefined;
    if (!recentCancellation(registrations).recent) return undefined;
    if (idRecheckCountRef.current >= 10) return undefined;
    const timer = window.setInterval(() => {
      idRecheckCountRef.current++;
      if (idRecheckCountRef.current > 10) { window.clearInterval(timer); return; }
      reloadRegistrationsForIdCheck().catch(() => { /* */ });
    }, 30000);
    return () => window.clearInterval(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, registrations, reloadRegistrationsForIdCheck]);

  // v11.70: kein Modal mehr beim Event-Öffnen — der Hinweis steht ab
  // jetzt direkt als Box oben in der Teilnehmerliste, solange die
  // Bedingung erfüllt ist (siehe Render-Block unten). Der Ref bleibt
  // erhalten, um in Zukunft ein erneutes „Mount-Trigger"-Verhalten
  // einbauen zu können, ohne den Save-Pfad zu touchen.
  React.useEffect(() => {
    if (!selectedEvent || isLoadingRegs || registrations.length === 0) return;
    if (idFixCheckedForRef.current === selectedEvent.id) return;
    idFixCheckedForRef.current = selectedEvent.id;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, registrations, isLoadingRegs]);

  // v6.17: Spaltenkonfiguration der Teilnehmertabelle (pro Event, lokal gespeichert).
  //  - columnOrder = geordnete Liste sichtbarer Spalten-IDs
  //  - hiddenColumns = ausgeblendete Spalten-IDs (können übers "+ Spalte"-Popover wieder zugeschaltet werden)
  // Die Spezialspalten 'id' / 'vorname' / 'nachname' / 'action' sind alwaysVisible und können nicht ausgeblendet werden.
  const [columnOrder, setColumnOrder] = React.useState<string[]>([]);
  const [hiddenColumns, setHiddenColumns] = React.useState<string[]>([]);
  const [showColumnPicker, setShowColumnPicker] = React.useState(false);
  // v26.44: „Matches anzeigen" — gruppiert die Teilnehmer-Tabelle in gegenseitige
  // Roommate-Paare (Match 1, Match 2, …) + Rest-Cluster. Nur relevant, wenn das
  // Event überhaupt eine Roommate-Spalte hat.
  const [showMatches, setShowMatches] = React.useState(false);
  return {
    columnOrder, getActiveTeams, hiddenColumns, moveRegToTeam, onTeamDrop, openTeamMailDialog,
    sendTeamMails, setColumnOrder, setHiddenColumns, setShowColumnPicker, setShowMatches,
    showColumnPicker, showMatches,
  };
}

