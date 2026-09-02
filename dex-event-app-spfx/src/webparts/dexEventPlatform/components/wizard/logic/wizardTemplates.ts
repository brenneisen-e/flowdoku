import * as React from 'react';
import { CustomFieldInput } from '../../wizard/customFieldInput';
import { SubEventDraft } from '../../wizard/wizardTypes';
import { AgendaItem } from '../../../types';

/* applyTemplate — aus EventCreationPage.tsx ausgelagert (Zeilen 2691-2753 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface ApplyTemplateCtx {
  b2runStartblocks: string[];
  isDe: boolean;
  setB2runStartblocks: React.Dispatch<React.SetStateAction<string[]>>;
  setCustomFields: React.Dispatch<React.SetStateAction<CustomFieldInput[]>>;
  setSelectedTemplate: React.Dispatch<React.SetStateAction<"blank" | "b2run">>;
}

export function applyTemplateImpl(ctx: ApplyTemplateCtx, template: "blank" | "b2run"): void {
  const { b2runStartblocks, isDe, setB2runStartblocks, setCustomFields, setSelectedTemplate } = ctx;
    setSelectedTemplate(template);
    if (template === 'blank') {
      // v7.20-Fix: NICHT alle Fields löschen — nur die B2Run-spezifischen
      // (Präfix "b2run_"). So gehen Custom-Felder, die der Organizer manuell
      // angelegt hat, beim Deselect des B2Run-Templates nicht verloren.
      setCustomFields(prev => prev.filter(f => !f.id.startsWith('b2run_')));
      setB2runStartblocks([]);
      return;
    }
    if (template === 'b2run') {
      // Custom Fields in der Reihenfolge der B2Run-Excel-Spalten
      // Hinweis: Strasse/PLZ/Stadt werden NICHT abgefragt (werden leer in der Excel stehen)
      // Locale-abhängige Labels/Optionen. IDs bleiben konstant, damit die
      // B2Run-Logik (Infoservice -> Mobilnummer, CSV-Export etc.) unabhängig
      // von der Sprache funktioniert.
      const fields: CustomFieldInput[] = isDe ? [
        { id: 'b2run_startblock', label: 'Startblock', type: 'select', required: true, options: [...b2runStartblocks], visible: true },
        { id: 'b2run_gruppe', label: 'Gruppe', type: 'select', required: true, options: ['offene Klasse', 'Nordic Walker', 'Damen', 'Herren'], visible: true },
        { id: 'b2run_altersklasse', label: 'Altersklasse', type: 'select', required: true, options: ['unter 18', '18-29', '30-39', '40-49', '50-59', '60+'], visible: true },
        { id: 'b2run_infoservice', label: 'Infoservice nutzen (SMS von B2Run — Mobilnummer erforderlich)', type: 'checkbox', required: false, options: [], visible: true },
        { id: 'b2run_mobilnummer', label: 'Mobilnummer (nur bei aktiviertem Infoservice)', type: 'text', required: true, options: [], visible: true, showIf: { fieldId: 'b2run_infoservice', values: ['true'] } },
        { id: 'b2run_anonym', label: 'Anonym teilnehmen', type: 'checkbox', required: false, options: [], visible: true },
        { id: 'b2run_laufshirt', label: 'Deloitte-Laufshirt', type: 'select', required: true, options: ['Habe bereits ein Laufshirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true },
        {
          id: 'b2run_datenschutz',
          label: 'Zustimmung AGB, Datenschutz & Bildaufnahmen',
          type: 'checkbox',
          required: true,
          options: [],
          visible: true,
          externalLinks: [
            { label: 'AGB (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
            { label: 'Datenschutz (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
          ],
        },
      ] : [
        { id: 'b2run_startblock', label: 'Start block', type: 'select', required: true, options: [...b2runStartblocks], visible: true },
        { id: 'b2run_gruppe', label: 'Category', type: 'select', required: true, options: ['Open class', 'Nordic Walker', 'Women', 'Men'], visible: true },
        { id: 'b2run_altersklasse', label: 'Age group', type: 'select', required: true, options: ['under 18', '18-29', '30-39', '40-49', '50-59', '60+'], visible: true },
        { id: 'b2run_infoservice', label: 'Use B2Run info service (SMS — mobile number required)', type: 'checkbox', required: false, options: [], visible: true },
        { id: 'b2run_mobilnummer', label: 'Mobile number (only if info service is enabled)', type: 'text', required: true, options: [], visible: true, showIf: { fieldId: 'b2run_infoservice', values: ['true'] } },
        { id: 'b2run_anonym', label: 'Participate anonymously', type: 'checkbox', required: false, options: [], visible: true },
        { id: 'b2run_laufshirt', label: 'Deloitte running shirt', type: 'select', required: true, options: ['I already have one', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true },
        {
          id: 'b2run_datenschutz',
          label: 'I agree to the terms, privacy policy and photo/video recordings',
          type: 'checkbox',
          required: true,
          options: [],
          visible: true,
          externalLinks: [
            { label: 'Terms (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
            { label: 'Privacy (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
          ],
        },
      ];
      // v7.20-Fix: bestehende NON-b2run-Felder erhalten und die B2Run-Felder
      // anhängen (vorher: setCustomFields(fields) hat alles überschrieben).
      setCustomFields(prev => {
        const nonB2run = prev.filter(f => !f.id.startsWith('b2run_'));
        return [...nonB2run, ...fields];
      });
    }
}

/* resetDemoVariantBaseState — aus EventCreationPage.tsx ausgelagert (Zeilen 2831-2864 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface ResetDemoVariantBaseStateCtx {
  setAgenda: React.Dispatch<React.SetStateAction<AgendaItem[]>>;
  setAskSalutation: React.Dispatch<React.SetStateAction<boolean>>;
  setAskTeamName: React.Dispatch<React.SetStateAction<boolean>>;
  setAudience: React.Dispatch<React.SetStateAction<string>>;
  setConfirmDialogEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  setConfirmDialogMode: React.Dispatch<React.SetStateAction<string>>;
  setConfirmDialogText: React.Dispatch<React.SetStateAction<string>>;
  setContactEmail: React.Dispatch<React.SetStateAction<string>>;
  setContactInfo: React.Dispatch<React.SetStateAction<string>>;
  setContactName: React.Dispatch<React.SetStateAction<string>>;
  setCustomFields: React.Dispatch<React.SetStateAction<CustomFieldInput[]>>;
  setDurchstarterCapacity: React.Dispatch<React.SetStateAction<string>>;
  setEmailLanguage: React.Dispatch<React.SetStateAction<string>>;
  setEventImageUrl: React.Dispatch<React.SetStateAction<string>>;
  setFunstarterCapacity: React.Dispatch<React.SetStateAction<string>>;
  setKlammerDeadline: React.Dispatch<React.SetStateAction<string>>;
  setLocationFilter: React.Dispatch<React.SetStateAction<string>>;
  setNoDescription: React.Dispatch<React.SetStateAction<boolean>>;
  setRemovedSavedSubs: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setSplitLabelA: React.Dispatch<React.SetStateAction<string>>;
  setSplitLabelB: React.Dispatch<React.SetStateAction<string>>;
  setSplitSharedWaitlist: React.Dispatch<React.SetStateAction<boolean>>;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setTeamJoinRequiresApproval: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamOpenSlotsVisible: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamPartialAllowed: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamRegistrationEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamSize: React.Dispatch<React.SetStateAction<number>>;
  setTransferTimes: React.Dispatch<React.SetStateAction<{ id: string; location: string; meetingPoint: string; address: string; date: string; departureTime: string; arrivalTime: string; description: string; }[]>>;
  setUseSplitCapacities: React.Dispatch<React.SetStateAction<boolean>>;
  setWaitlistEnabled: React.Dispatch<React.SetStateAction<boolean>>;
}

export function resetDemoVariantBaseStateImpl(ctx: ResetDemoVariantBaseStateCtx): void {
  const { setAgenda, setAskSalutation, setAskTeamName, setAudience, setConfirmDialogEnabled, setConfirmDialogMode, setConfirmDialogText, setContactEmail, setContactInfo, setContactName, setCustomFields, setDurchstarterCapacity, setEmailLanguage, setEventImageUrl, setFunstarterCapacity, setKlammerDeadline, setLocationFilter, setNoDescription, setRemovedSavedSubs, setSplitLabelA, setSplitLabelB, setSplitSharedWaitlist, setSubEvents, setTeamJoinRequiresApproval, setTeamOpenSlotsVisible, setTeamPartialAllowed, setTeamRegistrationEnabled, setTeamSize, setTransferTimes, setUseSplitCapacities, setWaitlistEnabled } = ctx;
    // v28.7: Demo-Vorlagen setzen immer eine Beschreibung — der
    // „Keine Beschreibung"-Schalter darf dann nicht angehakt bleiben.
    setNoDescription(false);
    // v28.20: keine Klammer-Frist aus einer vorherigen Vorlage mitschleppen.
    setKlammerDeadline('');
    setUseSplitCapacities(false);
    setSplitLabelA('Teilnehmergruppe 1');
    setSplitLabelB('Teilnehmergruppe 2');
    setDurchstarterCapacity('0');
    setFunstarterCapacity('0');
    setSplitSharedWaitlist(false);
    setTeamRegistrationEnabled(false);
    setTeamSize(4);
    setAskTeamName(false);
    setTeamPartialAllowed(false);
    setTeamOpenSlotsVisible(false);
    setTeamJoinRequiresApproval(false);
    setAskSalutation(false);
    setConfirmDialogEnabled(false); // v18.75: Sicherheitshinweis-Default
    setConfirmDialogMode('summary');
    setConfirmDialogText('');
    setSubEvents([]);
    setRemovedSavedSubs([]);
    setCustomFields([]);
    setAgenda([]);
    setTransferTimes([]);
    setLocationFilter('');
    setAudience('');
    setEventImageUrl('');
    setContactName('');
    setContactEmail('');
    setContactInfo('');
    setWaitlistEnabled(true);
    setEmailLanguage('DE');
}

/* applyEventTemplate — aus EventCreationPage.tsx ausgelagert (Zeilen 2872-2964 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface ApplyEventTemplateCtx {
  resetDemoVariantBaseState: () => void;
  setAddrCity: React.Dispatch<React.SetStateAction<string>>;
  setAddrHouseNo: React.Dispatch<React.SetStateAction<string>>;
  setAddrStreet: React.Dispatch<React.SetStateAction<string>>;
  setAddrZip: React.Dispatch<React.SetStateAction<string>>;
  setAgenda: React.Dispatch<React.SetStateAction<AgendaItem[]>>;
  setAskSalutation: React.Dispatch<React.SetStateAction<boolean>>;
  setAudience: React.Dispatch<React.SetStateAction<string>>;
  setBilingualFields: React.Dispatch<React.SetStateAction<boolean>>;
  setCurrentStep: React.Dispatch<React.SetStateAction<number>>;
  setCustomFields: React.Dispatch<React.SetStateAction<CustomFieldInput[]>>;
  setDescription: React.Dispatch<React.SetStateAction<string>>;
  setDurchstarterCapacity: React.Dispatch<React.SetStateAction<string>>;
  setEmailLanguage: React.Dispatch<React.SetStateAction<string>>;
  setExcludedUsers: React.Dispatch<React.SetStateAction<string[]>>;
  setFilterMode: React.Dispatch<React.SetStateAction<"AND" | "OR">>;
  setFunstarterCapacity: React.Dispatch<React.SetStateAction<string>>;
  setImageFile: React.Dispatch<React.SetStateAction<File>>;
  setImageOrigAspect: React.Dispatch<React.SetStateAction<number>>;
  setImageOrigFile: React.Dispatch<React.SetStateAction<File>>;
  setImagePreview: React.Dispatch<React.SetStateAction<string>>;
  setLocation: React.Dispatch<React.SetStateAction<string>>;
  setLocationFilter: React.Dispatch<React.SetStateAction<string>>;
  setMaxParticipants: React.Dispatch<React.SetStateAction<string>>;
  setShowTemplatePicker: React.Dispatch<React.SetStateAction<boolean>>;
  setSplitDescA: React.Dispatch<React.SetStateAction<string>>;
  setSplitDescB: React.Dispatch<React.SetStateAction<string>>;
  setSplitHelpText: React.Dispatch<React.SetStateAction<string>>;
  setSplitLabelA: React.Dispatch<React.SetStateAction<string>>;
  setSplitLabelB: React.Dispatch<React.SetStateAction<string>>;
  setSplitSectionTitle: React.Dispatch<React.SetStateAction<string>>;
  setSplitSharedWaitlist: React.Dispatch<React.SetStateAction<boolean>>;
  setTemplateLoadingId: React.Dispatch<React.SetStateAction<string>>;
  setTitle: React.Dispatch<React.SetStateAction<string>>;
  setUnlimitedParticipants: React.Dispatch<React.SetStateAction<boolean>>;
  setUseSplitCapacities: React.Dispatch<React.SetStateAction<boolean>>;
  setWaitlistEnabled: React.Dispatch<React.SetStateAction<boolean>>;
}

export async function applyEventTemplateImpl(ctx: ApplyEventTemplateCtx, ev: import("../../../types/index").DeloitteEvent): Promise<void> {
  const { resetDemoVariantBaseState, setAddrCity, setAddrHouseNo, setAddrStreet, setAddrZip, setAgenda, setAskSalutation, setAudience, setBilingualFields, setCurrentStep, setCustomFields, setDescription, setDurchstarterCapacity, setEmailLanguage, setExcludedUsers, setFilterMode, setFunstarterCapacity, setImageFile, setImageOrigAspect, setImageOrigFile, setImagePreview, setLocation, setLocationFilter, setMaxParticipants, setShowTemplatePicker, setSplitDescA, setSplitDescB, setSplitHelpText, setSplitLabelA, setSplitLabelB, setSplitSectionTitle, setSplitSharedWaitlist, setTemplateLoadingId, setTitle, setUnlimitedParticipants, setUseSplitCapacities, setWaitlistEnabled } = ctx;
    setTemplateLoadingId(ev.id);
    try {
      resetDemoVariantBaseState();
      setTitle(`${ev.title || ''} (Kopie)`.trim());
      setDescription(ev.description || '');
      setLocation(ev.location || '');
      setAddrStreet(ev.locationAddress?.street || '');
      setAddrHouseNo(ev.locationAddress?.houseNo || '');
      setAddrZip(ev.locationAddress?.zip || '');
      setAddrCity(ev.locationAddress?.city || '');
      setLocationFilter((ev.locationAudience || []).join(', '));
      setAudience((ev.audienceFilter || []).join(', '));
      setFilterMode(ev.filterMode || 'OR');
      if (ev.maxParticipants && ev.maxParticipants > 0) { setUnlimitedParticipants(false); setMaxParticipants(String(ev.maxParticipants)); }
      else { setUnlimitedParticipants(true); setMaxParticipants(''); }
      setWaitlistEnabled(!!ev.waitlistEnabled);
      setAskSalutation(!!ev.askSalutation);
      if (ev.agenda && ev.agenda.length > 0) setAgenda([...ev.agenda]);
      // Geteilte Kapazität übernehmen, falls vorhanden.
      if ((ev.splitLabelA || '').trim() || (ev.splitLabelB || '').trim() || (ev.durchstarterCapacity || 0) > 0 || (ev.funstarterCapacity || 0) > 0) {
        setUseSplitCapacities(true);
        setSplitLabelA(ev.splitLabelA || '');
        setSplitLabelB(ev.splitLabelB || '');
        setSplitDescA(ev.splitDescA || '');
        setSplitDescB(ev.splitDescB || '');
        setSplitHelpText(ev.splitHelpText || '');
        setSplitSectionTitle(ev.splitSectionTitle || '');
        setDurchstarterCapacity(String(ev.durchstarterCapacity || 0));
        setFunstarterCapacity(String(ev.funstarterCapacity || 0));
        setSplitSharedWaitlist(!!ev.splitSharedWaitlist);
      }
      setCustomFields((ev.eventSpecificFields || []).map(f => ({
        id: f.id, label: f.label, type: f.type, required: f.required,
        options: f.options ? [...f.options] : [], visible: true,
        ...(f.multi ? { multi: true } : {}),
        ...(f.helpText ? { helpText: f.helpText } : {}),
        ...(f.helpTextStyle === 'inline' ? { helpTextStyle: 'inline' as const } : {}),
        ...(f.showIf ? { showIf: { fieldId: f.showIf.fieldId, values: [...f.showIf.values] } } : {}),
        ...(f.onlyForGroup ? { onlyForGroup: f.onlyForGroup } : {}),
        ...(f.confirmLabel ? { confirmLabel: f.confirmLabel } : {}),
        ...(f.labelEn ? { labelEn: f.labelEn } : {}),
        ...(f.helpTextEn ? { helpTextEn: f.helpTextEn } : {}),
        ...(f.confirmLabelEn ? { confirmLabelEn: f.confirmLabelEn } : {}),
        ...(f.optionsEn && f.optionsEn.length > 0 ? { optionsEn: [...f.optionsEn] } : {}),
        ...(f.externalLinks && f.externalLinks.length > 0 ? { externalLinks: f.externalLinks.map(x => ({ ...x })) } : {}),
        ...(f.ccOnEmails ? { ccOnEmails: true } : {}),
        ...(f.notifyRoommate === false ? { notifyRoommate: false } : {}),
        ...(f.audienceOnly ? { audienceOnly: true } : {}),
        // v29.20 (Audit A3): auch hier fehlten Vorauswahl, Vorfilter,
        // Uhrzeit-Option und die daterange-Grenzen — die Kopie verlor sie.
        ...(f.defaultValue ? { defaultValue: f.defaultValue } : {}),
        ...(f.optionCategories && f.optionCategories.length > 0 ? { optionCategories: [...f.optionCategories] } : {}),
        ...(f.prefilterLabel ? { prefilterLabel: f.prefilterLabel } : {}),
        ...(f.withTime ? { withTime: true } : {}),
        ...(f.rangeStart ? { rangeStart: f.rangeStart } : {}),
        ...(f.rangeEnd ? { rangeEnd: f.rangeEnd } : {}),
        ...(typeof f.maxNights === 'number' && f.maxNights > 0 ? { maxNights: f.maxNights } : {}),
      })));
      // v29.20 (Audit): Die EN-Varianten oben nützen nur mit dem Schalter —
      // serializeCustomFields schreibt sie ausschließlich bei aktivem
      // bilingualFields. Ohne die Übernahme zeigte der Wizard die EN-Texte
      // der Vorlage an und der Save warf sie still weg. Ebenso übernimmt die
      // Kopie jetzt Mail-Sprache und Ausschluss-Liste — beides gehört zur
      // Konfiguration, die man mit einer Vorlage erwartet.
      setBilingualFields(!!ev.bilingualFields);
      setEmailLanguage((ev.emailLanguage || 'DE').toUpperCase() === 'EN' ? 'EN' : 'DE');
      setExcludedUsers([...(ev.excludedUsers || [])]);
      // Bild: Vorschau sofort, Datei best-effort vom SP-Anhang nachladen.
      // v29.20 (Audit): Vorher NUR gesetzt, wenn die Vorlage ein Bild hat —
      // beim Wechsel von Vorlage A (mit Bild) zu Vorlage B (ohne) blieben
      // imagePreview/imageFile von A stehen, und das neue Event bekam still
      // das Foto des falschen Events. resetDemoVariantBaseState leerte nur
      // eventImageUrl. Jetzt wird immer erst geleert.
      setImagePreview('');
      setImageFile(null);
      setImageOrigFile(null);
      setImageOrigAspect(null);
      if (ev.imageUrl) {
        setImagePreview(ev.imageUrl);
        try {
          const resp = await fetch(ev.imageUrl);
          if (resp.ok) {
            const blob = await resp.blob();
            const ext = (blob.type && blob.type.indexOf('png') >= 0) ? 'png' : 'jpg';
            setImageFile(new File([blob], `vorlage-bild.${ext}`, { type: blob.type || 'image/jpeg' }));
          }
        } catch { /* nur Vorschau, kein Re-Upload */ }
      }
      setShowTemplatePicker(false);
      setCurrentStep(0);
    } finally {
      setTemplateLoadingId(null);
    }
}

/* loadDemoSubEvent — aus EventCreationPage.tsx ausgelagert (Zeilen 3036-3073 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface LoadDemoSubEventCtx {
  beforeNextSaturday: (daysBefore: number, hour: number, minute: number) => Date;
  berlinLocalToUtcIso: (localStr: string) => string;
  fmtDatetime: (d: Date) => string;
  nextSaturdayAt: (hour: number, minute: number) => Date;
  resetDemoVariantBaseState: () => void;
  setAskSalutation: React.Dispatch<React.SetStateAction<boolean>>;
  setCurrentStep: React.Dispatch<React.SetStateAction<number>>;
  setCustomFields: React.Dispatch<React.SetStateAction<CustomFieldInput[]>>;
  setDescription: React.Dispatch<React.SetStateAction<string>>;
  setEndDate: React.Dispatch<React.SetStateAction<string>>;
  setLastDeregisterDate: React.Dispatch<React.SetStateAction<string>>;
  setLocation: React.Dispatch<React.SetStateAction<string>>;
  setMaxParticipants: React.Dispatch<React.SetStateAction<string>>;
  setRegistrationDeadline: React.Dispatch<React.SetStateAction<string>>;
  setStartDate: React.Dispatch<React.SetStateAction<string>>;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setTitle: React.Dispatch<React.SetStateAction<string>>;
  setUseSplitCapacities: React.Dispatch<React.SetStateAction<boolean>>;
  setWaitlistEnabled: React.Dispatch<React.SetStateAction<boolean>>;
}

export function loadDemoSubEventImpl(ctx: LoadDemoSubEventCtx): void {
  const { beforeNextSaturday, berlinLocalToUtcIso, fmtDatetime, nextSaturdayAt, resetDemoVariantBaseState, setAskSalutation, setCurrentStep, setCustomFields, setDescription, setEndDate, setLastDeregisterDate, setLocation, setMaxParticipants, setRegistrationDeadline, setStartDate, setSubEvents, setTitle, setUseSplitCapacities, setWaitlistEnabled } = ctx;
    resetDemoVariantBaseState();
    const start = nextSaturdayAt(9, 0);
    const end = nextSaturdayAt(17, 0);
    const deadline = beforeNextSaturday(3, 23, 59);
    setTitle('Demo-Conference mit Dinner');
    setDescription('Hauptkonferenz + abendliches Dinner als getrenntes Sub-Event mit eigener Anmeldung.');
    setLocation('Deloitte Office Hamburg');
    setStartDate(fmtDatetime(start));
    setEndDate(fmtDatetime(end));
    // v30.67: fmtDatetime — der Fristen-State ist 16-stellig (s. loadDemoStandard).
    setRegistrationDeadline(fmtDatetime(deadline));
    setLastDeregisterDate(fmtDatetime(deadline));
    setMaxParticipants('100');
    setUseSplitCapacities(false);
    setWaitlistEnabled(true);
    setAskSalutation(false);
    const tDemo = Date.now();
    setCustomFields([
      { id: `cf-${tDemo}`, label: 'Hotel-Buchung', type: 'select', required: false,
        options: ['Ja, ich brauche ein Hotel', 'Nein, ich reise abends ab'], visible: true },
    ]);
    const dinnerStart = nextSaturdayAt(18, 0);
    const dinnerEnd = nextSaturdayAt(22, 0);
    setSubEvents([
      {
        id: `se-${tDemo}`,
        title: 'Networking-Dinner',
        startDate: berlinLocalToUtcIso(fmtDatetime(dinnerStart)),
        endDate: berlinLocalToUtcIso(fmtDatetime(dinnerEnd)),
        registrationDeadline: '',
        location: 'Restaurant Fischmarkt',
        description: 'Optionales Networking-Dinner im Anschluss an die Konferenz.',
        maxParticipants: 60,
        disableEmails: false,
        disableOutlook: false,
        customFields: [],
      },
    ]);
    setCurrentStep(0);
}

/* loadDemoSubEventTeam — aus EventCreationPage.tsx ausgelagert (Zeilen 3077-3120 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface LoadDemoSubEventTeamCtx {
  beforeNextSaturday: (daysBefore: number, hour: number, minute: number) => Date;
  berlinLocalToUtcIso: (localStr: string) => string;
  fmtDatetime: (d: Date) => string;
  nextSaturdayAt: (hour: number, minute: number) => Date;
  resetDemoVariantBaseState: () => void;
  setAskSalutation: React.Dispatch<React.SetStateAction<boolean>>;
  setAskTeamName: React.Dispatch<React.SetStateAction<boolean>>;
  setCurrentStep: React.Dispatch<React.SetStateAction<number>>;
  setCustomFields: React.Dispatch<React.SetStateAction<CustomFieldInput[]>>;
  setDescription: React.Dispatch<React.SetStateAction<string>>;
  setEndDate: React.Dispatch<React.SetStateAction<string>>;
  setLastDeregisterDate: React.Dispatch<React.SetStateAction<string>>;
  setLocation: React.Dispatch<React.SetStateAction<string>>;
  setMaxParticipants: React.Dispatch<React.SetStateAction<string>>;
  setRegistrationDeadline: React.Dispatch<React.SetStateAction<string>>;
  setStartDate: React.Dispatch<React.SetStateAction<string>>;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setTeamJoinRequiresApproval: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamOpenSlotsVisible: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamPartialAllowed: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamRegistrationEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamSize: React.Dispatch<React.SetStateAction<number>>;
  setTitle: React.Dispatch<React.SetStateAction<string>>;
  setUseSplitCapacities: React.Dispatch<React.SetStateAction<boolean>>;
  setWaitlistEnabled: React.Dispatch<React.SetStateAction<boolean>>;
}

export function loadDemoSubEventTeamImpl(ctx: LoadDemoSubEventTeamCtx): void {
  const { beforeNextSaturday, berlinLocalToUtcIso, fmtDatetime, nextSaturdayAt, resetDemoVariantBaseState, setAskSalutation, setAskTeamName, setCurrentStep, setCustomFields, setDescription, setEndDate, setLastDeregisterDate, setLocation, setMaxParticipants, setRegistrationDeadline, setStartDate, setSubEvents, setTeamJoinRequiresApproval, setTeamOpenSlotsVisible, setTeamPartialAllowed, setTeamRegistrationEnabled, setTeamSize, setTitle, setUseSplitCapacities, setWaitlistEnabled } = ctx;
    resetDemoVariantBaseState();
    const start = nextSaturdayAt(18, 0);
    const end = nextSaturdayAt(22, 0);
    const deadline = beforeNextSaturday(5, 23, 59);
    setTitle('Demo-Kneipenquiz mit Team-Anmeldung');
    setDescription('Quizabend, bei dem ganze Teams über das Anmeldeformular angemeldet werden.');
    setLocation('Heinrich Campus Düsseldorf, 6. Etage, Dachterrasse');
    setStartDate(fmtDatetime(start));
    setEndDate(fmtDatetime(end));
    setRegistrationDeadline(fmtDatetime(deadline)); // v30.67: s. loadDemoStandard
    setLastDeregisterDate(fmtDatetime(deadline));
    setMaxParticipants('80');
    setUseSplitCapacities(false);
    setWaitlistEnabled(true);
    setAskSalutation(false);
    setTeamRegistrationEnabled(true);
    setTeamSize(4);
    setAskTeamName(true);
    setTeamPartialAllowed(true);
    setTeamOpenSlotsVisible(true);
    setTeamJoinRequiresApproval(false);
    const tDemo = Date.now();
    setCustomFields([
      { id: `cf-${tDemo}`, label: 'Essenspräferenz', type: 'select', required: true,
        options: ['Vegetarisch', 'Vegan', 'Keine Einschränkungen'], visible: true },
    ]);
    const briefStart = nextSaturdayAt(17, 0);
    const briefEnd = nextSaturdayAt(17, 30);
    setSubEvents([
      {
        id: `se-${tDemo}`,
        title: 'Vorbereitungs-Briefing (Quizmaster)',
        startDate: berlinLocalToUtcIso(fmtDatetime(briefStart)),
        endDate: berlinLocalToUtcIso(fmtDatetime(briefEnd)),
        registrationDeadline: '',
        location: 'Heinrich Campus Düsseldorf, 6. Etage, Dachterrasse',
        description: 'Kurzes Briefing für die Quizmaster-Helfer vor dem Event.',
        maxParticipants: 10,
        disableEmails: false,
        disableOutlook: false,
        customFields: [],
      },
    ]);
    setCurrentStep(0);
}

