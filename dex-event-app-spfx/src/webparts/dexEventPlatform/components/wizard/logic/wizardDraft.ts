import * as React from 'react';
import { CustomFieldInput } from '../../wizard/customFieldInput';
import { AgendaItem } from '../../../types';
import { SubEventDraft } from '../../wizard/wizardTypes';
import { EmailOverrideEntry } from '../../wizard/emailOverrideEntry';

/* applyDraftPayload — aus EventCreationPage.tsx ausgelagert (Zeilen 4314-4372 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface ApplyDraftPayloadCtx {
  canBilling: boolean;
  setActiveFrom: React.Dispatch<React.SetStateAction<string>>;
  setAddrCity: React.Dispatch<React.SetStateAction<string>>;
  setAddrHouseNo: React.Dispatch<React.SetStateAction<string>>;
  setAddrStreet: React.Dispatch<React.SetStateAction<string>>;
  setAddrZip: React.Dispatch<React.SetStateAction<string>>;
  setAgenda: React.Dispatch<React.SetStateAction<AgendaItem[]>>;
  setAskSalutation: React.Dispatch<React.SetStateAction<boolean>>;
  setAskTeamName: React.Dispatch<React.SetStateAction<boolean>>;
  setAudience: React.Dispatch<React.SetStateAction<string>>;
  setBillingFields: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  setBillingRelevant: React.Dispatch<React.SetStateAction<boolean>>;
  setBillingSendMode: React.Dispatch<React.SetStateAction<"manual" | "auto">>;
  setCancelRuleAfter: React.Dispatch<React.SetStateAction<boolean>>;
  setCancelRuleAmount: React.Dispatch<React.SetStateAction<number>>;
  setCancelRuleEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  setCancelRuleUnit: React.Dispatch<React.SetStateAction<"days" | "hours">>;
  setContactEmail: React.Dispatch<React.SetStateAction<string>>;
  setContactInfo: React.Dispatch<React.SetStateAction<string>>;
  setContactName: React.Dispatch<React.SetStateAction<string>>;
  setCurrentStep: React.Dispatch<React.SetStateAction<number>>;
  setCustomFields: React.Dispatch<React.SetStateAction<CustomFieldInput[]>>;
  setDescription: React.Dispatch<React.SetStateAction<string>>;
  setDisableEmails: React.Dispatch<React.SetStateAction<boolean>>;
  setDisableOutlook: React.Dispatch<React.SetStateAction<boolean>>;
  setEmailTemplateOverrides: React.Dispatch<React.SetStateAction<Record<string, EmailOverrideEntry>>>;
  setEndDate: React.Dispatch<React.SetStateAction<string>>;
  setExcludedUsers: React.Dispatch<React.SetStateAction<string[]>>;
  setFilterMode: React.Dispatch<React.SetStateAction<"AND" | "OR">>;
  setKlammerDeadline: React.Dispatch<React.SetStateAction<string>>;
  setLastDeregisterDate: React.Dispatch<React.SetStateAction<string>>;
  setLocation: React.Dispatch<React.SetStateAction<string>>;
  setLocationFilter: React.Dispatch<React.SetStateAction<string>>;
  setMaxParticipants: React.Dispatch<React.SetStateAction<string>>;
  setNoCancelAfterDeadline: React.Dispatch<React.SetStateAction<boolean>>;
  setOnlineMeetingMode: React.Dispatch<React.SetStateAction<"none" | "own" | "auto">>;
  setOpenRuleDays: React.Dispatch<React.SetStateAction<number>>;
  setOpenRuleEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  setOpenRuleFixedDate: React.Dispatch<React.SetStateAction<string>>;
  setOpenRuleMode: React.Dispatch<React.SetStateAction<"day" | "week">>;
  setOrganizer: React.Dispatch<React.SetStateAction<string>>;
  setOrganizerEmails: React.Dispatch<React.SetStateAction<string[]>>;
  setRegistrationDeadline: React.Dispatch<React.SetStateAction<string>>;
  setRegRuleAmount: React.Dispatch<React.SetStateAction<number>>;
  setRegRuleEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  setRegRuleUnit: React.Dispatch<React.SetStateAction<"days" | "hours">>;
  setRequireSubEventSelection: React.Dispatch<React.SetStateAction<boolean>>;
  setStartDate: React.Dispatch<React.SetStateAction<string>>;
  setSubEventCalendar: React.Dispatch<React.SetStateAction<boolean>>;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setSubEventSingleChoice: React.Dispatch<React.SetStateAction<boolean>>;
  setSubEventsOnlyMode: React.Dispatch<React.SetStateAction<boolean>>;
  setSubEventsOptIn: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamRegistrationEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamSize: React.Dispatch<React.SetStateAction<number>>;
  setTeamsLink: React.Dispatch<React.SetStateAction<string>>;
  setTitle: React.Dispatch<React.SetStateAction<string>>;
  setUserCancelAllowed: React.Dispatch<React.SetStateAction<boolean>>;
  setVisAllSubs: React.Dispatch<React.SetStateAction<boolean>>;
  setWaitlistEnabled: React.Dispatch<React.SetStateAction<boolean>>;
}

export function applyDraftPayloadImpl(ctx: ApplyDraftPayloadCtx, d: Record<string, unknown>): void {
  const { canBilling, setActiveFrom, setAddrCity, setAddrHouseNo, setAddrStreet, setAddrZip, setAgenda, setAskSalutation, setAskTeamName, setAudience, setBillingFields, setBillingRelevant, setBillingSendMode, setCancelRuleAfter, setCancelRuleAmount, setCancelRuleEnabled, setCancelRuleUnit, setContactEmail, setContactInfo, setContactName, setCurrentStep, setCustomFields, setDescription, setDisableEmails, setDisableOutlook, setEmailTemplateOverrides, setEndDate, setExcludedUsers, setFilterMode, setKlammerDeadline, setLastDeregisterDate, setLocation, setLocationFilter, setMaxParticipants, setNoCancelAfterDeadline, setOnlineMeetingMode, setOpenRuleDays, setOpenRuleEnabled, setOpenRuleFixedDate, setOpenRuleMode, setOrganizer, setOrganizerEmails, setRegistrationDeadline, setRegRuleAmount, setRegRuleEnabled, setRegRuleUnit, setRequireSubEventSelection, setStartDate, setSubEventCalendar, setSubEvents, setSubEventSingleChoice, setSubEventsOnlyMode, setSubEventsOptIn, setTeamRegistrationEnabled, setTeamSize, setTeamsLink, setTitle, setUserCancelAllowed, setVisAllSubs, setWaitlistEnabled } = ctx;
    const str = (v: unknown): string => (typeof v === 'string' ? v : '');
    const bool = (v: unknown, dflt: boolean): boolean => (typeof v === 'boolean' ? v : dflt);
    const num = (v: unknown, dflt: number): number => (typeof v === 'number' && isFinite(v) ? v : dflt);
    setTitle(str(d.title)); setDescription(str(d.description)); setLocation(str(d.location));
    setAddrStreet(str(d.addrStreet)); setAddrHouseNo(str(d.addrHouseNo)); setAddrZip(str(d.addrZip)); setAddrCity(str(d.addrCity));
    setOrganizer(str(d.organizer));
    if (Array.isArray(d.organizerEmails)) setOrganizerEmails(d.organizerEmails as string[]);
    setContactName(str(d.contactName)); setContactEmail(str(d.contactEmail)); setContactInfo(str(d.contactInfo));
    setStartDate(str(d.startDate)); setEndDate(str(d.endDate));
    setRegistrationDeadline(str(d.registrationDeadline)); setLastDeregisterDate(str(d.lastDeregisterDate));
    setKlammerDeadline(str(d.klammerDeadline)); setActiveFrom(str(d.activeFrom));
    setMaxParticipants(str(d.maxParticipants)); setWaitlistEnabled(bool(d.waitlistEnabled, false));
    setAudience(str(d.audience)); setLocationFilter(str(d.locationFilter));
    setFilterMode(d.filterMode === 'AND' ? 'AND' : 'OR');
    if (Array.isArray(d.excludedUsers)) setExcludedUsers(d.excludedUsers as string[]);
    if (Array.isArray(d.customFields)) setCustomFields(d.customFields as CustomFieldInput[]);
    if (Array.isArray(d.agenda)) setAgenda(d.agenda as AgendaItem[]);
    if (Array.isArray(d.subEvents)) setSubEvents((d.subEvents as SubEventDraft[]).map(x => ({ ...x, imageFile: null })));
    setSubEventsOptIn(bool(d.subEventsOptIn, false));
    setSubEventsOnlyMode(bool(d.subEventsOnlyMode, false));
    setSubEventCalendar(bool(d.subEventCalendar, false));
    setSubEventSingleChoice(bool(d.subEventSingleChoice, false));
    setRequireSubEventSelection(bool(d.requireSubEventSelection, false));
    setAskSalutation(bool(d.askSalutation, false));
    setTeamRegistrationEnabled(bool(d.teamRegistrationEnabled, false));
    setTeamSize(num(d.teamSize, 2)); setAskTeamName(bool(d.askTeamName, false));
    setUserCancelAllowed(bool(d.userCancelAllowed, true));
    setNoCancelAfterDeadline(bool(d.noCancelAfterDeadline, false));
    setTeamsLink(str(d.teamsLink));
    // v30.26: Modus aus dem Entwurf; Alt-Entwürfe kennen ihn nicht — dort
    // ergibt sich 'own' aus einem vorhandenen Link, sonst 'none'.
    const draftOmMode = str(d.onlineMeetingMode);
    setOnlineMeetingMode(
      draftOmMode === 'auto' || draftOmMode === 'own' || draftOmMode === 'none'
        ? draftOmMode
        : (str(d.teamsLink).trim() ? 'own' : 'none'),
    );
    setDisableEmails(bool(d.disableEmails, false)); setDisableOutlook(bool(d.disableOutlook, false));
    if (d.emailTemplateOverrides && typeof d.emailTemplateOverrides === 'object') {
      setEmailTemplateOverrides(d.emailTemplateOverrides as Record<string, EmailOverrideEntry>);
    }
    setOpenRuleEnabled(bool(d.openRuleEnabled, false));
    setOpenRuleMode(d.openRuleMode === 'week' ? 'week' : 'day');
    setOpenRuleDays(num(d.openRuleDays, 7)); setOpenRuleFixedDate(str(d.openRuleFixedDate));
    setRegRuleEnabled(bool(d.regRuleEnabled, false)); setRegRuleAmount(num(d.regRuleAmount, 1));
    setRegRuleUnit(d.regRuleUnit === 'hours' ? 'hours' : 'days');
    setCancelRuleEnabled(bool(d.cancelRuleEnabled, false)); setCancelRuleAmount(num(d.cancelRuleAmount, 1));
    setCancelRuleUnit(d.cancelRuleUnit === 'hours' ? 'hours' : 'days');
    setCancelRuleAfter(bool(d.cancelRuleAfter, false));
    setVisAllSubs(bool(d.visAllSubs, false));
    if (typeof d.billingRelevant === 'boolean') setBillingRelevant(d.billingRelevant);
    setBillingSendMode(d.billingSendMode === 'auto' ? 'auto' : 'manual');
    if (d.billingFields && typeof d.billingFields === 'object') setBillingFields(d.billingFields as Record<string, string>);
    // v30.59: Auch beim Entwurf nach OBEN klemmen. Ein Entwurf, der auf
    // Schritt 10 gespeichert wurde (als Admin, oder mit einem Build, in dem es
    // den Schritt noch für alle gab), führte beim Wieder-Öffnen als Organizer
    // auf einen Schritt, den es dort nicht gibt — und der Assistent zeigte
    // gar nichts mehr an. Dasselbe Muster wie beim Tour-Schritt oben.
    setCurrentStep(Math.min(Math.max(0, num(d.currentStep, 0)), canBilling ? 9 : 8));
}

