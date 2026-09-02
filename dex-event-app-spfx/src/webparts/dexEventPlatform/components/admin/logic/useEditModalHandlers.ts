/* useEditModalHandlers — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 1265-1482 des
 * Stands vor dem Schnitt). Der Rumpf ist zeichengleich uebernommen; was die
 * Gruppe aus dem Komponenten-Scope liest, kommt als `ctx` herein, was sie
 * nach aussen liefert, geht als Objekt zurueck.
 */
import * as React from 'react';
import { EventService, REG_LIST_NAME, SPRegistration } from '../../../services/EventService';
import { isDeloitteInternalEmail } from '../../../utils/deloitteDomain';
import { DeloitteEvent } from '../../../types';

export interface UseEditModalHandlersCtx {
  currentUser: import("../../../types/index").User;
  editForm: Record<string, string>;
  editingReg: SPRegistration;
  eventServiceRef: EventService;
  getAllRegistrations: (eventId: string, onHttpError?: (_status: number) => void) => Promise<SPRegistration[]>;
  isDe: boolean;
  searchUser: (email: string) => Promise<{ displayName: string; location: string; jobTitle: string; department?: string; mobilePhone?: string; company?: string; }>;
  selectedEvent: DeloitteEvent;
  setEditError: React.Dispatch<React.SetStateAction<string>>;
  setEditForm: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  setEditingReg: React.Dispatch<React.SetStateAction<SPRegistration>>;
  setIsSavingEdit: React.Dispatch<React.SetStateAction<boolean>>;
  setRegistrations: React.Dispatch<React.SetStateAction<SPRegistration[]>>;
}

export interface UseEditModalHandlersResult {
  closeEditModal: () => void;
  openEditModal: (reg: SPRegistration) => void;
  saveEdit: () => Promise<void>;
}

export function useEditModalHandlers(ctx: UseEditModalHandlersCtx): UseEditModalHandlersResult {
  const {
    currentUser, editForm, editingReg, eventServiceRef, getAllRegistrations, isDe, searchUser,
    selectedEvent, setEditError, setEditForm, setEditingReg, setIsSavingEdit, setRegistrations,
  } = ctx;
  const openEditModal = (reg: SPRegistration): void => {
    setEditError('');
    setEditingReg(reg);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const r = reg as any;
    const initial: Record<string, string> = {
      Anrede: r.Anrede || '',
      Vorname: r.Vorname || '',
      Nachname: r.Nachname || '',
      ParticipantEmail: r.ParticipantEmail || '',
      Phone: r.Phone || '',
      Department: r.Department || '',
      Location: r.Location || '',
      JobTitle: r.JobTitle || '',
      Status: r.Status || '',
      // v10.13+: B2Run-Felder ins Edit-Form aufnehmen damit das B2Run-Modul
      // im Edit-Modal die aktuellen Werte vorbefüllen kann. Strings (auch
      // wenn leer) — bei Nicht-B2Run-Events werden die Felder im Modal
      // sowieso nicht angezeigt.
      StarterType: r.StarterType || '',
      PreferredStarterType: r.PreferredStarterType || '',
    };
    // Custom-Field-Werte aus dem reg laden (sie sind als SP-Spalten gespeichert)
    if (selectedEvent?.eventSpecificFields) {
      for (const f of selectedEvent.eventSpecificFields) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sp = (f as any).spInternalName || '';
        if (sp) initial[sp] = (r[sp] !== undefined && r[sp] !== null) ? String(r[sp]) : '';
      }
    }
    setEditForm(initial);
  };
  const closeEditModal = (): void => {
    setEditingReg(null);
    setEditForm({});
    setEditError('');
  };
  const saveEdit = async (): Promise<void> => {
    if (!editingReg || !eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setIsSavingEdit(true);
    setEditError('');
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const r = editingReg as any;

      // Stamm-Daten (Vorname, Nachname, E-Mail) werden seit v9.7 ebenfalls
      // editierbar gemacht — z.B. um Tippfehler nach manueller Anlage zu
      // korrigieren. Validierung:
      //   1. E-Mail muss eine Deloitte-Deutschland-Adresse sein (@deloitte.de).
      //      Die Plattform ist nur für DEALL freigeschaltet — auch @deloitte.com
      //      (US/Global) zählt als extern. Sonst Abbruch mit Fehler.
      //   2. Person muss in M365 existieren (searchUserByEmail). Sonst
      //      Abbruch mit "Tippfehler"-Hinweis.
      // Die uebrigen Profil-Felder (Phone, Department, Location, JobTitle)
      // bleiben read-only — sie kommen aus dem M365-Profil.
      const oldVorname = String(r.Vorname || '');
      const oldNachname = String(r.Nachname || '');
      const oldEmail = String(r.ParticipantEmail || '');
      const newVorname = (editForm.Vorname || '').trim();
      const newNachname = (editForm.Nachname || '').trim();
      const newEmail = (editForm.ParticipantEmail || '').trim();
      const stammChanged = newVorname !== oldVorname || newNachname !== oldNachname || newEmail !== oldEmail;

      const profileFields: { Department?: string; Location?: string; JobTitle?: string } = {};
      if (stammChanged) {
        // Plausibilität: nicht-leer
        if (!newVorname || !newNachname || !newEmail) {
          setEditError(isDe
            ? 'Vorname, Nachname und E-Mail dürfen nicht leer sein.'
            : 'First name, last name and email must not be empty.');
          return;
        }
        // Domain-Check: nur Deloitte-Adressen zulassen (v27.11: beliebige
        // Member-Firm-Domain, konsistent zur International-Suche v26.57).
        const lower = newEmail.toLowerCase();
        const isDeloitte = isDeloitteInternalEmail(lower);
        if (!isDeloitte) {
          setEditError(isDe
            ? `Externe E-Mail-Adresse — nicht erlaubt. Bitte eine Deloitte-Adresse verwenden (z.B. @deloitte.de oder eine andere Member-Firm-Domain).`
            : `External email address — not allowed. Please use a Deloitte mailbox (e.g. @deloitte.de or another member-firm domain).`);
          return;
        }
        // Existenz-Check via M365-Profile (UPN!=SMTP-aware). Wenn wir hier
        // nichts finden, ist es entweder ein Tippfehler oder ein Account
        // der gar nicht (mehr) im Tenant ist — beides nicht akzeptabel.
        if (newEmail.toLowerCase() !== oldEmail.toLowerCase()) {
          const profile = await searchUser(newEmail);
          if (!profile || !profile.displayName) {
            setEditError(isDe
              ? `Person mit E-Mail "${newEmail}" wurde im Deloitte-Tenant nicht gefunden. Bitte Adresse prüfen (Tippfehler?).`
              : `No person found in the Deloitte tenant for "${newEmail}". Please check the address (typo?).`);
            return;
          }
          // Profil-Daten gleich mit-übernehmen, damit der Eintrag konsistent
          // bleibt (Department / Location / JobTitle passen zum neuen User).
          profileFields.Department = ''; // searchUser liefert displayName/location/jobTitle
          profileFields.Location = profile.location || '';
          profileFields.JobTitle = profile.jobTitle || '';
        }
      }

      const oldValues: Record<string, unknown> = {};
      const patch: Record<string, unknown> = {};
      const fieldLabelMap: Record<string, string> = {};

      if (stammChanged) {
        if (newVorname !== oldVorname) {
          oldValues.Vorname = oldVorname; patch.Vorname = newVorname; fieldLabelMap.Vorname = isDe ? 'Vorname' : 'First name';
        }
        if (newNachname !== oldNachname) {
          oldValues.Nachname = oldNachname; patch.Nachname = newNachname; fieldLabelMap.Nachname = isDe ? 'Nachname' : 'Last name';
        }
        if (newEmail !== oldEmail) {
          oldValues.ParticipantEmail = oldEmail; patch.ParticipantEmail = newEmail; fieldLabelMap.ParticipantEmail = 'E-Mail';
          // Profil-Daten mit aktualisieren (nur wenn überhaupt was zurückkam)
          if (profileFields.Location) {
            oldValues.Location = String(r.Location || ''); patch.Location = profileFields.Location;
            fieldLabelMap.Location = isDe ? 'Standort' : 'Location';
          }
          if (profileFields.JobTitle) {
            oldValues.JobTitle = String(r.JobTitle || ''); patch.JobTitle = profileFields.JobTitle;
            fieldLabelMap.JobTitle = 'Job Title';
          }
        }
      }

      // Custom-Felder des Events. v10.15+: nur Felder ins Patch aufnehmen die
      // sich tatsächlich geändert haben — sonst sendet ein unverändertes
      // Choice-Feld ohne ausgewählten Wert einen leeren String an SP, der
      // mit HTTP 400 'Invalid choice' kippt und das ganze Update abbricht.
      if (selectedEvent?.eventSpecificFields) {
        for (const f of selectedEvent.eventSpecificFields) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const sp = (f as any).spInternalName || '';
          if (!sp) continue;
          const oldVal = r[sp] !== undefined && r[sp] !== null ? String(r[sp]) : '';
          const newVal = editForm[sp] || '';
          if (newVal === oldVal) continue;  // unverändert → skip
          fieldLabelMap[sp] = f.label;
          oldValues[sp] = oldVal;
          patch[sp] = newVal;
        }
      }

      // v10.13+: B2Run-Felder explizit ins Patch aufnehmen, wenn sich was
      // geändert hat. Sind keine regulären customFields, daher werden sie
      // oben in der eventSpecificFields-Loop nicht abgeholt. Nur bei
      // Split-Capacity-Events relevant.
      const isSplitEvent = !!selectedEvent
        && (selectedEvent.durchstarterCapacity || 0) > 0
        && (selectedEvent.funstarterCapacity || 0) > 0;
      if (isSplitEvent) {
        const oldStarter = String(r.StarterType || '');
        const newStarter = (editForm.StarterType || '').trim();
        if (newStarter !== oldStarter) {
          oldValues.StarterType = oldStarter;
          patch.StarterType = newStarter;
          fieldLabelMap.StarterType = isDe ? 'Starter-Typ' : 'Starter type';
        }
        const oldPref = String(r.PreferredStarterType || '');
        const newPref = (editForm.PreferredStarterType || '').trim();
        if (newPref !== oldPref) {
          oldValues.PreferredStarterType = oldPref;
          patch.PreferredStarterType = newPref;
          fieldLabelMap.PreferredStarterType = isDe ? 'Wunsch-Starter-Typ' : 'Preferred starter type';
        }
      }
      if (Object.keys(patch).length === 0) {
        // Keine Änderung — nichts zu tun.
        closeEditModal();
        return;
      }
      const actor = {
        name: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email,
        email: currentUser.email,
      };
      const ok = await eventServiceRef.adminUpdateRegistration(
        selectedEvent.subsiteUrl, editingReg.Id, patch, actor, oldValues, fieldLabelMap
      );
      if (!ok) {
        // Häufigste 400-Ursache: eine SP-Spalte aus dem Patch existiert nicht
        // auf dieser Teilnehmerliste (z.B. StarterType auf einem v9-Event ohne
        // B2Run-Schema, oder ein neu hinzugefügtes Custom-Field ohne 'Spalten
        // fixen'-Run). Hilfreicher Hinweis auf den Repair-Button.
        setEditError(isDe
          ? 'Speichern fehlgeschlagen — vermutlich fehlt eine SP-Spalte in der Teilnehmerliste. Klicke einmal „Spalten fixen" im Toolbox-Bereich des Events, dann erneut versuchen.'
          : 'Save failed — likely a missing SP column on the participant list. Click „Fix columns" in the event toolbox once, then retry.');
        return;
      }
      // v9.0: Audit-Log mit Diff der geänderten Felder
      try {
        const changes: Record<string, { old: unknown; new: unknown }> = {};
        for (const k of Object.keys(patch)) {
          if (oldValues[k] !== patch[k]) changes[k] = { old: oldValues[k], new: patch[k] };
        }
        await eventServiceRef.writeChangeLog({
          action: 'ParticipantUpdated',
          targetType: 'Participant',
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          targetId: ((editingReg as any).ParticipantEmail || '') + '#' + editingReg.Id,
          targetName: `${editingReg.Vorname || ''} ${editingReg.Nachname || ''}`.trim(),
          eventId: selectedEvent.id,
          eventTitle: selectedEvent.title,
          details: { changes },
        });
      } catch { /* */ }
      // v30.67: Die E-Mail ist der Schlüssel des Registers DEX_Participants
      // (`$filter=Email eq …`) — und der einzige Schlüssel des Datenmodells
      // überhaupt. Jeder Pfad, der ihn VERGIBT, schreibt doppelt (Zeile +
      // Register); dieser Pfad, der ihn ÄNDERT, behandelte ihn wie ein
      // Textfeld. Folge: Das Event blieb unter der alten Adresse im Register,
      // „Meine Events" der Person war leer, die Doppel-Anmelde-Prüfung fragte
      // unter der neuen Adresse ins Leere, und beim Abmelden fand
      // `removeParticipantEvent` nichts — ein verwaister Verweis für immer.
      // Reihenfolge: prüfbare Nebenbuchhaltung zuerst (neu anlegen), dann alt
      // entfernen. Schlägt das Anlegen fehl, erfährt es der Organizer.
      let registryFailed = false;
      if (patch.ParticipantEmail !== undefined && newEmail.toLowerCase() !== oldEmail.toLowerCase()) {
        const rowStatus = String(r.Status || '');
        const evNo = selectedEvent.eventNumber || 0;
        if (evNo > 0 && rowStatus !== 'Abgemeldet') {
          const regStatus = rowStatus === 'Warteliste' ? 'Warteliste' : 'Angemeldet';
          const up = await eventServiceRef.upsertParticipant(newVorname, newNachname, newEmail, evNo, regStatus);
          if (up) {
            const rm = await eventServiceRef.removeParticipantEvent(oldEmail, evNo);
            // false heißt hier auch „unter der alten Adresse gab es keinen
            // Eintrag" — nicht unterscheidbar, deshalb nur ein Warn-Log; die
            // Register-Prüfung (Admin) findet einen Rest als verwaisten Verweis.
            if (!rm) console.warn('[DEX] saveEdit: Register-Eintrag der alten Adresse nicht entfernt:', oldEmail, evNo);
          } else {
            registryFailed = true;
          }
        }
        // Die Zeile gehört jetzt der neuen Person — mit ReadSecurity=2 sieht
        // sie ihre eigene Zeile sonst nicht (Zeilen-Autor bleibt die alte).
        try { await eventServiceRef.trySetItemAuthor(selectedEvent.subsiteUrl, REG_LIST_NAME, editingReg.Id, newEmail); } catch { /* best-effort, s. trySetItemAuthor */ }
      }
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
      if (registryFailed) {
        setEditError(isDe
          ? 'Die Zeile ist gespeichert, aber das Teilnehmer-Register (DEX_Participants) konnte nicht auf die neue Adresse umgeschrieben werden. Die Person sieht das Event unter „Meine Events“ erst, wenn ein Admin das Register abgleicht. Bitte noch einmal speichern oder den Admin informieren.'
          : 'The row was saved, but the participant registry (DEX_Participants) could not be switched to the new address. The person will not see the event under „My events“ until an admin reconciles the registry. Please save again or inform the admin.');
        return;
      }
      closeEditModal();
    } catch (err) {
      console.warn('[DEX] saveEdit error:', err);
      setEditError(isDe
        ? 'Unerwarteter Fehler beim Speichern.'
        : 'Unexpected error while saving.');
    } finally {
      setIsSavingEdit(false);
    }
  };
  return {
    closeEditModal, openEditModal, saveEdit,
  };
}

