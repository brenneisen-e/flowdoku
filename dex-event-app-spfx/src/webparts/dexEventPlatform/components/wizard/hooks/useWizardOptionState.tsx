/* useWizardOptionState — aus EventCreationPage.tsx ausgelagert (Zeilen 1979-2431 des
 * damaligen Stands). Der Bereich ist zeichengleich uebernommen, inklusive der
 * Reihenfolge seiner Hook-Aufrufe: der Aufruf dieses Hooks steht in der Seite
 * an genau der Stelle, an der der Bereich vorher stand. Schwerpunkt: Optionen (Upload, Team, Benachrichtigungen, Quiz).
 *
 * Die Grenze ist mechanisch gezogen (zusammenhaengender Bereich), nicht
 * thematisch — der Name beschreibt den Schwerpunkt, nicht eine reine Trennung. */
import * as React from 'react';
import { DeloitteEvent } from '../../../types';
import { SubEventDraft, SuggestedEntry } from '../../wizard/wizardTypes';
import { getSuggestedFieldsCatalog } from '../../../data/suggestedFields';
import { CustomFieldInput } from '../../wizard/customFieldInput';
import { renderShowIfConfigImpl } from '../../wizard/logic/wizardRenderHelpers';
import { applyEventTemplateImpl, applyTemplateImpl, resetDemoVariantBaseStateImpl } from '../../wizard/logic/wizardTemplates';
import { AgendaItem } from '../../../types';

export interface UseWizardOptionStateCtx {
  b2runStartblocks: string[];
  customFields: CustomFieldInput[];
  durchstarterCapacity: string;
  editEvent: import("../../../types/index").DeloitteEvent;
  funstarterCapacity: string;
  isDe: boolean;
  isEditMode: boolean;
  newStartblock: string;
  selectedTemplate: "blank" | "b2run";
  setAddrCity: React.Dispatch<React.SetStateAction<string>>;
  setAddrHouseNo: React.Dispatch<React.SetStateAction<string>>;
  setAddrStreet: React.Dispatch<React.SetStateAction<string>>;
  setAddrZip: React.Dispatch<React.SetStateAction<string>>;
  setAgenda: React.Dispatch<React.SetStateAction<AgendaItem[]>>;
  setAudience: React.Dispatch<React.SetStateAction<string>>;
  setB2runStartblocks: React.Dispatch<React.SetStateAction<string[]>>;
  setContactEmail: React.Dispatch<React.SetStateAction<string>>;
  setContactInfo: React.Dispatch<React.SetStateAction<string>>;
  setContactName: React.Dispatch<React.SetStateAction<string>>;
  setCurrentStep: React.Dispatch<React.SetStateAction<number>>;
  setCustomFields: React.Dispatch<React.SetStateAction<CustomFieldInput[]>>;
  setDescription: React.Dispatch<React.SetStateAction<string>>;
  setDurchstarterCapacity: React.Dispatch<React.SetStateAction<string>>;
  setEmailLanguage: React.Dispatch<React.SetStateAction<string>>;
  setEndDate: React.Dispatch<React.SetStateAction<string>>;
  setEventImageUrl: React.Dispatch<React.SetStateAction<string>>;
  setExcludedUsers: React.Dispatch<React.SetStateAction<string[]>>;
  setFieldExpandOverride: React.Dispatch<React.SetStateAction<Record<string, boolean>>>;
  setFilterMode: React.Dispatch<React.SetStateAction<"AND" | "OR">>;
  setFunstarterCapacity: React.Dispatch<React.SetStateAction<string>>;
  setImageFile: React.Dispatch<React.SetStateAction<File>>;
  setImageOrigAspect: React.Dispatch<React.SetStateAction<number>>;
  setImageOrigFile: React.Dispatch<React.SetStateAction<File>>;
  setImagePreview: React.Dispatch<React.SetStateAction<string>>;
  setKlammerDeadline: React.Dispatch<React.SetStateAction<string>>;
  setLastDeregisterDate: React.Dispatch<React.SetStateAction<string>>;
  setLocation: React.Dispatch<React.SetStateAction<string>>;
  setLocationFilter: React.Dispatch<React.SetStateAction<string>>;
  setMaxParticipants: React.Dispatch<React.SetStateAction<string>>;
  setNewStartblock: React.Dispatch<React.SetStateAction<string>>;
  setNoDescription: React.Dispatch<React.SetStateAction<boolean>>;
  setRegistrationDeadline: React.Dispatch<React.SetStateAction<string>>;
  setRemovedSavedSubs: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setSelectedTemplate: React.Dispatch<React.SetStateAction<"blank" | "b2run">>;
  setShowTemplatePicker: React.Dispatch<React.SetStateAction<boolean>>;
  setSplitDescA: React.Dispatch<React.SetStateAction<string>>;
  setSplitDescB: React.Dispatch<React.SetStateAction<string>>;
  setSplitHelpText: React.Dispatch<React.SetStateAction<string>>;
  setSplitLabelA: React.Dispatch<React.SetStateAction<string>>;
  setSplitLabelB: React.Dispatch<React.SetStateAction<string>>;
  setSplitSectionTitle: React.Dispatch<React.SetStateAction<string>>;
  setSplitSharedWaitlist: React.Dispatch<React.SetStateAction<boolean>>;
  setStartDate: React.Dispatch<React.SetStateAction<string>>;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setTemplateLoadingId: React.Dispatch<React.SetStateAction<string>>;
  setTitle: React.Dispatch<React.SetStateAction<string>>;
  setTransferTimes: React.Dispatch<React.SetStateAction<{ id: string; location: string; meetingPoint: string; address: string; date: string; departureTime: string; arrivalTime: string; description: string; }[]>>;
  setUnlimitedParticipants: React.Dispatch<React.SetStateAction<boolean>>;
  setWaitlistEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
}

export function useWizardOptionState(ctx: UseWizardOptionStateCtx) {
  const { b2runStartblocks, customFields, durchstarterCapacity, editEvent, funstarterCapacity, isDe, isEditMode, newStartblock, selectedTemplate, setAddrCity, setAddrHouseNo, setAddrStreet, setAddrZip, setAgenda, setAudience, setB2runStartblocks, setContactEmail, setContactInfo, setContactName, setCurrentStep, setCustomFields, setDescription, setDurchstarterCapacity, setEmailLanguage, setEndDate, setEventImageUrl, setExcludedUsers, setFieldExpandOverride, setFilterMode, setFunstarterCapacity, setImageFile, setImageOrigAspect, setImageOrigFile, setImagePreview, setKlammerDeadline, setLastDeregisterDate, setLocation, setLocationFilter, setMaxParticipants, setNewStartblock, setNoDescription, setRegistrationDeadline, setRemovedSavedSubs, setSelectedTemplate, setShowTemplatePicker, setSplitDescA, setSplitDescB, setSplitHelpText, setSplitLabelA, setSplitLabelB, setSplitSectionTitle, setSplitSharedWaitlist, setStartDate, setSubEvents, setTemplateLoadingId, setTitle, setTransferTimes, setUnlimitedParticipants, setWaitlistEnabled, showAlert } = ctx;
  const [allowAttendeeUpload, setAllowAttendeeUpload] = React.useState<boolean>(
    !!editEvent?.allowAttendeeUpload
  );
  const [attendeeUploadHint, setAttendeeUploadHint] = React.useState<string>(
    editEvent?.attendeeUploadHint || ''
  );
  const [attendeeUploadLabel, setAttendeeUploadLabel] = React.useState<string>(
    editEvent?.attendeeUploadLabel || ''
  );
  // v11.80: Anrede im Registrierungsformular abfragen (Default false). Wenn
  // false, wird das Anrede-Dropdown ausgeblendet und ein leerer String als
  // Anrede gespeichert. Wird im neuen Schritt 5 (Felder) konfiguriert.
  const [askSalutation, setAskSalutation] = React.useState<boolean>(
    !!editEvent?.askSalutation
  );
  // v18.75: Sicherheitshinweis vor dem Absenden der Anmeldung (Schritt 5, ganz
  // unten). Default aus. Modus 'summary' = Auswahl-Übersicht (Haupt-/Sub-Events
  // mit De-/Selektieren), 'freetext' = eigener Hinweis-Text.
  const [confirmDialogEnabled, setConfirmDialogEnabled] = React.useState<boolean>(!!editEvent?.confirmDialogEnabled);
  const [confirmDialogMode, setConfirmDialogMode] = React.useState<string>(editEvent?.confirmDialogMode || 'summary');
  const [confirmDialogText, setConfirmDialogText] = React.useState<string>(editEvent?.confirmDialogText || '');
  // v18.35: Anmeldesprache vorgeben. '' = App-Sprache (Default), 'de' / 'en' =
  // Anmeldeseite (inkl. Disclaimer) immer in dieser Sprache anzeigen.
  const [registrationLanguage, setRegistrationLanguage] = React.useState<'' | 'de' | 'en'>(
    editEvent?.registrationLanguage === 'de' || editEvent?.registrationLanguage === 'en' ? editEvent.registrationLanguage : ''
  );
  // v20.2: Self-Check-in-States aus dem Wizard entfernt — Aktivierung läuft
  // automatisch beim ersten Klick auf die Aktionen (Check-in-Seite, Admin
  // Center, QR-Kachel im Event-Detail); Zeitfenster + Deaktivieren im
  // Kachel-Modal des Admin Centers. Der Wizard fasst die SelfCheckIn*-Spalten
  // weder beim Create noch beim Edit an, damit die dort gesetzten Werte einen
  // Wizard-Save überleben.
  // v11.80: Team-Anmeldung — eine Person meldet ein ganzes Team an.
  // Konfiguration im neuen Schritt 4 (Team-Anmeldung). Die tatsächliche
  // Multi-Person-Anmelde-Logik folgt mit v11.81+; aktuell wird nur die
  // Konfiguration persistiert.
  const [teamRegistrationEnabled, setTeamRegistrationEnabled] = React.useState<boolean>(
    !!editEvent?.teamRegistrationEnabled
  );
  const [teamSize, setTeamSize] = React.useState<number>(
    typeof editEvent?.teamSize === 'number' && editEvent.teamSize > 0 ? editEvent.teamSize : 4
  );
  const [askTeamName, setAskTeamName] = React.useState<boolean>(
    !!editEvent?.askTeamName
  );
  // v11.81: Erweiterte Team-Konfiguration — Beitritts-Modus, Sichtbarkeit
  // offener Slots, Lead-Approval. Die tatsächliche Team-Anmelde-Logik
  // (Multi-Person-Form, Mails, Outlook) folgt mit v11.82+.
  const [teamPartialAllowed, setTeamPartialAllowed] = React.useState<boolean>(
    !!editEvent?.teamPartialAllowed
  );
  const [teamOpenSlotsVisible, setTeamOpenSlotsVisible] = React.useState<boolean>(
    !!editEvent?.teamOpenSlotsVisible
  );
  const [teamJoinRequiresApproval, setTeamJoinRequiresApproval] = React.useState<boolean>(
    !!editEvent?.teamJoinRequiresApproval
  );
  // v22.78: frei benennbarer Team-Begriff (z.B. „Break-Out Session") +
  // „Teilnehmer dürfen keine neuen Teams erstellen".
  const [teamTermSingular, setTeamTermSingular] = React.useState<string>(editEvent?.teamTermSingular || '');
  const [teamTermPlural, setTeamTermPlural] = React.useState<string>(editEvent?.teamTermPlural || '');
  const [teamMembersCannotCreate, setTeamMembersCannotCreate] = React.useState<boolean>(!!editEvent?.teamMembersCannotCreate);
  // v24.58: Anzeige-Bezeichnung des Haupt-Events in der Sub-Event-Auswahl.
  const [mainEventLabelMode, setMainEventLabelMode] = React.useState<'default' | 'custom' | 'none'>(editEvent?.mainEventLabelMode || 'default');
  const [mainEventLabel, setMainEventLabel] = React.useState<string>(editEvent?.mainEventLabel || '');
  // v17.20: Bilingual-Toggle — wenn an, kann der Organizer pro Custom-Field
  // (Label, Help-Text, Checkbox-Confirm-Text, Dropdown-Optionen) eine
  // englische Variante hinterlegen. Wird im Wizard-Schritt 5 ganz oben als
  // separater Toggle eingestellt; die EN-Inputs blenden pro Card auf, wenn
  // der Toggle aktiv ist.
  const [bilingualFields, setBilingualFields] = React.useState<boolean>(
    !!editEvent?.bilingualFields
  );
  // v6.15: Starter-Typ → Startblock-Zuordnung + Leistungsnachweis-Pflicht
  const [durchstarterStartblock, setDurchstarterStartblock] = React.useState<string>(
    editEvent?.durchstarterStartblock || ''
  );
  const [funstarterStartblock, setFunstarterStartblock] = React.useState<string>(
    editEvent?.funstarterStartblock || ''
  );
  // v10.24: setDurchstarterRequiresProof wird nicht mehr aufgerufen — der UI-
  // Toggle in Schritt 3 ist entfallen, das Feature wird durch Pro-Gruppe-
  // Custom-Fields in Schritt 4 ersetzt. State bleibt erhalten, damit
  // bestehende Events mit gesetztem Wert nicht beim Save den Wert verlieren
  // (durchstarterRequiresProof wird beim Persist mitgeschrieben falls
  // editEvent das Flag schon hatte).
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const [durchstarterRequiresProof, setDurchstarterRequiresProof] = React.useState<boolean>(
    !!editEvent?.durchstarterRequiresProof
  );
  const [showPreview, setShowPreview] = React.useState(false);
  const [showRegisterPreview, setShowRegisterPreview] = React.useState(false);
  const [triedNext, setTriedNext] = React.useState(false);
  // v26.86: Die Blöcke im Schritt „Kapazität & Sichtbarkeit" sind einzeln
  // einklappbar und beim ersten Aufruf EINGEKLAPPT (Set enthält die offenen
  // Keys → leer = alles zu), damit der Schritt nicht überfordert. Der äußere
  // Block-<div> (inkl. zebraS3Bg-Alternation) bleibt erhalten; nur der Body
  // wird ein-/ausgeblendet, die Überschrift ist der Klappschalter.
  // Bei ausgelöster Validierung (triedNext) klappt automatisch ALLES auf, damit
  // keine Fehlermeldung in einem eingeklappten Block versteckt bleibt.
  const [expandedVisBlocks, setExpandedVisBlocks] = React.useState<Set<string>>(() => new Set());
  const isVisOpen = (k: string): boolean => triedNext || expandedVisBlocks.has(k);
  const toggleVis = (k: string): void => setExpandedVisBlocks(prev => { const n = new Set(prev); if (n.has(k)) n.delete(k); else n.add(k); return n; });
  const visHeader = (key: string, badge: React.ReactNode, title: React.ReactNode): React.ReactElement => (
    <button type="button" onClick={() => toggleVis(key)} aria-expanded={isVisOpen(key)}
      style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 8, background: 'none', border: 'none', padding: 0, marginBottom: isVisOpen(key) ? 8 : 0, cursor: 'pointer', textAlign: 'left' }}>
      {badge}
      <span className="form-label" style={{ margin: 0, fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: 8 }}>{title}</span>
      <span style={{ marginLeft: 'auto', color: 'var(--dex-gray-400)', fontSize: '0.85rem', transform: isVisOpen(key) ? 'rotate(90deg)' : 'none', transition: 'transform .15s' }}>▶</span>
    </button>
  );
  // v22.62/v22.63: Beim „Weiter"/Speichern fragt ein Modal, ob die geänderte
  // Klammer-/Hauptevent-Sichtbarkeit auf alle Sub-Events übernommen werden soll
  // — IMMER wenn die Klammer-Sichtbarkeit geändert/neu gesetzt wurde UND von
  // den Sub-Events abweicht. `visSnapshotRef` hält den zuletzt „abgehandelten"
  // Sichtbarkeits-Stand (Baseline beim Mount via Effekt unten).
  const [visCopyModalOpen, setVisCopyModalOpen] = React.useState(false);
  const visCopyPendingRef = React.useRef<(() => void) | null>(null);
  const visSnapshotRef = React.useRef<string | null>(null);
  const [previewSections, setPreviewSections] = React.useState<Array<{ id: string; label: string }>>([
    { id: 'event', label: 'Event-Karte' },
    { id: 'personal', label: 'Personal Information' },
    { id: 'specific', label: 'Event specific Information' },
    { id: 'actions', label: 'Buttons' },
  ]);
  const [dragSectionId, setDragSectionId] = React.useState<string | null>(null);
  const [dragOverSectionId, setDragOverSectionId] = React.useState<string | null>(null);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [progress, setProgress] = React.useState(0);
  const [progressLabel, setProgressLabel] = React.useState('');
  // v19.x: „Sichtbarkeit prüfen"-Modal-State (Verteiler-Cache, Testpersonen-
  // Suche) ist nach <AudiencePicker> gewandert.
  const [submitted, setSubmitted] = React.useState(false);
  const [error, setError] = React.useState('');
  const [imageUploadError, setImageUploadError] = React.useState('');

  const locationOptions = ['Berlin', 'Dresden', 'Düsseldorf', 'Frankfurt', 'Görlitz', 'Halle', 'Hamburg', 'Hannover', 'Köln', 'Leipzig', 'Magdeburg', 'Mannheim', 'München', 'Nürnberg', 'Stuttgart', 'Walldorf', 'All'];

  const addCustomField = (): void => {
    const newId = `cf-${Date.now()}`;
    setCustomFields([...customFields, {
      id: newId, label: '', type: 'text',
      required: false, options: [], visible: true,
    }]);
    // v18.55: neues Feld direkt aufgeklappt, damit man es sofort ausfüllen kann.
    setFieldExpandOverride(prev => ({ ...prev, [newId]: true }));
  };

  /**
   * Deloitte-Standard-Vorschläge als Katalog. Der Organizer wählt über ein
   * Modal mit Checkboxen aus, welche dieser Felder hinzugefügt werden sollen.
   * Ausgewählte Felder werden ans Ende der aktuellen customFields angehängt.
   */
  // Bilingual: Labels + Optionen der Felder werden in der Event-Sprache (DE/EN)
  // angelegt, passend zum Locale beim Klick auf 'Vorgeschlagene Felder'.
  // v10.21: Catalog mit Kategorien — 'general' (default ausgeklappt) und
  // 'b2run' (default eingeklappt). Damit ersetzt das Suggested-Modal den
  // alten Template-Dropdown: User wählt fokussiert die Felder, die er
  // wirklich braucht, statt einen B2Run-Block auf einmal aufzuziehen.
  // v10.23: jeder Suggested-Field-Eintrag hat ein Fluent-UI-Icon (visuelles
  // Erkennungsmerkmal in der Auswahl-Liste) und einen ausführlicheren
  // Tooltip-Text — der erklärt dem Organizer, was das Feld in der App
  // bewirkt, ohne dass er es erst hinzufügen muss.
  const SUGGESTED_FIELDS_CATALOG: SuggestedEntry[] = getSuggestedFieldsCatalog(isDe);

  const [showSuggestedModal, setShowSuggestedModal] = React.useState(false);
  const [suggestedSelection, setSuggestedSelection] = React.useState<Record<string, boolean>>({});
  // v10.21: B2Run-Sektion im Suggested-Modal default eingeklappt — die meisten
  // Organizer brauchen sie nicht; soll nicht visuell übernehmen.
  const [showB2runSuggested, setShowB2runSuggested] = React.useState(false);

  const openSuggestedModal = (): void => {
    // v9.17: Standard ist KEINS ausgewählt — User wählt aktiv aus, was er
    // wirklich braucht. Vorher waren alle vorgewählt, was zu unbeabsichtigt
    // viele übernommenen Feldern führte.
    setSuggestedSelection({});
    setShowSuggestedModal(true);
  };

  const addSelectedSuggestedFields = (): void => {
    const selected = SUGGESTED_FIELDS_CATALOG.filter(s => suggestedSelection[s.key]);
    if (selected.length === 0) { setShowSuggestedModal(false); return; }
    // v22.38: Sonder-Eintrag „Anrede" schaltet das Standard-Anrede-Feld an
    // (askSalutation-Flag, Pseudo-Zeile in der Feld-Liste) statt ein
    // Custom-Field anzulegen.
    if (selected.some(s => s.key === 'salutation')) setAskSalutation(true);
    const buildable = selected.filter(s => s.key !== 'salutation');
    if (buildable.length === 0) { setShowSuggestedModal(false); return; }
    const now = Date.now();
    const newFields: CustomFieldInput[] = buildable.map((s, i) => s.build(now + i));
    // v10.21: B2Run-Felder haben deterministische IDs (b2run_startblock etc.).
    // Wenn ein Feld mit gleicher ID schon im customFields-Array steht, skippen
    // wir es — sonst entstehen Duplikate, wenn der User das Modal mehrfach
    // öffnet. Allgemeine Felder (cf-<timestamp>) bekommen eindeutige IDs und
    // werden immer angehängt.
    const existingIds = new Set(customFields.map(f => f.id));
    const dedupedNewFields = newFields.filter(f => !existingIds.has(f.id));
    setCustomFields([...customFields, ...dedupedNewFields]);
    setShowSuggestedModal(false);
  };

  const removeCustomField = (id: string): void => {
    setCustomFields(customFields.filter(f => f.id !== id));
  };

  const updateCustomField = (id: string, updates: Partial<CustomFieldInput>): void => {
    setCustomFields(customFields.map(f => f.id === id ? { ...f, ...updates } : f));
  };
  // v24.16: Sichtbarkeitsbedingung (showIf) als wiederverwendbarer Helfer —
  // genutzt vom Hauptevent UND von Sub-Event-Feldern (vorher fehlte die UI bei
  // Sub-Events komplett, daher liessen sich Bedingungen dort nie setzen).
  const renderShowIfConfig = (field: CustomFieldInput, idx: number, allFields: CustomFieldInput[], onUpdate: (u: Partial<CustomFieldInput>) => void): React.ReactElement => {
    return renderShowIfConfigImpl({
      isDe, showAlert,
    }, field, idx, allFields, onUpdate);
  };


  // === Sub-Event Custom-Field Helpers (v10.11+) =============================
  // Per-Sub-Event Custom-Fields ersetzen die hardcoded Funstarter/Durchstarter-
  // Frage. Pattern parallel zu den Hauptevent-Helpers — operieren aber auf dem
  // `customFields[]` eines spezifischen SubEventDraft (nach Client-`id`
  // identifiziert). Funktional minimaler als die Hauptevent-Variante (kein
  // Suggested-Modal, kein showIf für v1), reicht aber für „Auswahlfrage pro
  // Sub-Event mit individuellem Label + Optionen".
  const addSubEventCustomField = (subEventId: string): void => {
    setSubEvents(prev => prev.map(se => se.id !== subEventId ? se : ({
      ...se,
      customFields: [...(se.customFields || []), {
        id: `cf_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
        label: '',
        type: 'select',
        required: false,
        options: ['', ''],
        visible: true,
      }],
    })));
  };
  const removeSubEventCustomField = (subEventId: string, fieldId: string): void => {
    setSubEvents(prev => prev.map(se => se.id !== subEventId ? se : ({
      ...se,
      customFields: (se.customFields || []).filter(f => f.id !== fieldId),
    })));
  };
  const updateSubEventCustomField = (subEventId: string, fieldId: string, updates: Partial<CustomFieldInput>): void => {
    setSubEvents(prev => prev.map(se => se.id !== subEventId ? se : ({
      ...se,
      customFields: (se.customFields || []).map(f => f.id === fieldId ? { ...f, ...updates } : f),
    })));
  };
  const copyParentFieldsToSubEvent = (subEventId: string): void => {
    // Dupliziert die Hauptevent-Felder ins Sub-Event mit frischen IDs (sonst
    // kollidieren Field-IDs zwischen Parent und Children, was bei Validierungs-
    // Logik und showIf-Refs zu Konflikten führen würde).
    const cloned: CustomFieldInput[] = customFields.map(f => ({
      ...f,
      id: `cf_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
      options: f.options.slice(),
      externalLinks: f.externalLinks ? f.externalLinks.map(x => ({ ...x })) : undefined,
      showIf: undefined,  // showIf-Refs würden auf Parent-Field-IDs zeigen, droppen
    }));
    setSubEvents(prev => prev.map(se => se.id !== subEventId ? se : ({
      ...se,
      customFields: cloned,
    })));
  };

  /**
   * Template-Auswahl: setzt EventType und Custom Fields automatisch.
   * B2Run: legt alle Pflichtfelder für die Anmeldung bei b2run.com an
   * (laut Excel "Deloitte_Teilnehmer_innen_B2Run_Koeln_2025_v4.xlsx").
   *
   * v10.21: Template-Dropdown im Wizard entfällt; B2Run-Felder werden über
   * das Suggested-Felder-Modal einzeln gewählt. Diese Funktion bleibt für
   * eventuelle programmatische Aufrufer (Edit-Modus, Migrations-Skripte)
   * erhalten — sie wird im aktuellen UI nicht mehr aufgerufen.
   */
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const applyTemplate = (template: 'blank' | 'b2run'): void => {
    return applyTemplateImpl({
      b2runStartblocks, isDe, setB2runStartblocks, setCustomFields, setSelectedTemplate,
    }, template);
  };

  // Startblöcke-Änderung direkt in das Custom Field übernehmen
  React.useEffect(() => {
    if (selectedTemplate !== 'b2run' && !(isEditMode && customFields.some(f => f.id === 'b2run_startblock'))) return;
    setCustomFields(prev => prev.map(f =>
      f.id === 'b2run_startblock' ? { ...f, options: [...b2runStartblocks] } : f
    ));
  }, [b2runStartblocks]);

  // Edit-Mode: Wenn das Event B2Run-Custom-Fields hat, Startblöcke aus dem Field laden
  React.useEffect(() => {
    if (!isEditMode) return;
    const sb = customFields.find(f => f.id === 'b2run_startblock');
    if (sb && b2runStartblocks.length === 0 && sb.options && sb.options.length > 0) {
      const parts = sb.options.map(s => s.trim()).filter(Boolean);
      setB2runStartblocks(parts);
      setSelectedTemplate('b2run');
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isEditMode]);

  const addStartblock = (): void => {
    const trimmed = newStartblock.trim();
    if (!trimmed) return;
    if (b2runStartblocks.indexOf(trimmed) >= 0) { setNewStartblock(''); return; }
    setB2runStartblocks([...b2runStartblocks, trimmed]);
    setNewStartblock('');
  };

  const removeStartblock = (block: string): void => {
    setB2runStartblocks(b2runStartblocks.filter(b => b !== block));
  };

  // Bei B2Run: maxParticipants automatisch aus Summe von Durchstarter + Funstarter berechnen
  const isB2runTemplate = selectedTemplate === 'b2run' || (isEditMode && customFields.some(f => f.id === 'b2run_startblock'));
  // Seit v6.5: explizite Checkbox in Schritt 3 ("Lauf-Event mit getrennten
  // Starter-Kapazitäten") statt versteckt über das Template gesteuert.
  // Initial-Wert: beim Edit aus vorhandenen Kapazitäten abgeleitet, bei neuem
  // Event true wenn B2Run-Template gewählt wurde.
  const [useSplitCapacities, setUseSplitCapacities] = React.useState<boolean>(() => {
    if (editEvent) {
      return typeof editEvent.durchstarterCapacity === 'number'
        && typeof editEvent.funstarterCapacity === 'number'
        && (editEvent.durchstarterCapacity > 0 || editEvent.funstarterCapacity > 0);
    }
    return selectedTemplate === 'b2run';
  });
  // Automatisch aktivieren wenn B2Run-Template nachträglich gewählt wird.
  React.useEffect(() => {
    if (!editEvent && selectedTemplate === 'b2run') setUseSplitCapacities(true);
  }, [selectedTemplate, editEvent]);

  React.useEffect(() => {
    if (!useSplitCapacities) return;
    const d = parseInt(durchstarterCapacity, 10) || 0;
    const f = parseInt(funstarterCapacity, 10) || 0;
    const sum = d + f;
    if (sum > 0) setMaxParticipants(String(sum));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [durchstarterCapacity, funstarterCapacity, useSplitCapacities]);

  // v11.88: Helpers für Datums-Formatierung — werden von allen Demo-
  // Varianten + dem alten fillDemo geteilt.
  const fmtDatetime = (d: Date): string => {
    const pad = (n: number): string => (n < 10 ? '0' : '') + n;
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
  };
  const fmtDate = (d: Date): string => {
    const pad = (n: number): string => (n < 10 ? '0' : '') + n;
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  };

  // v11.88: Reset-Helfer — setzt alle Team-, Split- und sonstigen
  // Variant-spezifischen Felder auf neutralen Default zurück, damit die
  // Demo-Varianten nicht versehentlich Zustand der vorigen Variante erben.
  const resetDemoVariantBaseState = (): void => {
    return resetDemoVariantBaseStateImpl({
      setAgenda, setAskSalutation, setAskTeamName, setAudience, setConfirmDialogEnabled, setConfirmDialogMode,
      setConfirmDialogText, setContactEmail, setContactInfo, setContactName, setCustomFields, setDurchstarterCapacity,
      setEmailLanguage, setEventImageUrl, setFunstarterCapacity, setKlammerDeadline, setLocationFilter, setNoDescription,
      setRemovedSavedSubs, setSplitLabelA, setSplitLabelB, setSplitSharedWaitlist, setSubEvents, setTeamJoinRequiresApproval,
      setTeamOpenSlotsVisible, setTeamPartialAllowed, setTeamRegistrationEnabled, setTeamSize, setTransferTimes, setUseSplitCapacities,
      setWaitlistEnabled,
    });
  };

  // v24.9 (E): bestehendes Event als Vorlage übernehmen — Einstellungen + Bild
  // ins neue Formular laden (KEINE Datumswerte, KEINE Sub-Events — die legt der
  // Organizer fürs neue Event frisch fest). Das Bild wird vom (gleichen
  // SharePoint-)Anhang gefetcht und als Datei für den Re-Upload übernommen.
  const applyEventTemplate = async (ev: DeloitteEvent): Promise<void> => {
    return await applyEventTemplateImpl({
      resetDemoVariantBaseState, setAddrCity, setAddrHouseNo, setAddrStreet, setAddrZip, setAgenda,
      setAskSalutation, setAudience, setBilingualFields, setCurrentStep, setCustomFields, setDescription,
      setDurchstarterCapacity, setEmailLanguage, setExcludedUsers, setFilterMode, setFunstarterCapacity, setImageFile,
      setImageOrigAspect, setImageOrigFile, setImagePreview, setLocation, setLocationFilter, setMaxParticipants,
      setShowTemplatePicker, setSplitDescA, setSplitDescB, setSplitHelpText, setSplitLabelA, setSplitLabelB,
      setSplitSectionTitle, setSplitSharedWaitlist, setTemplateLoadingId, setTitle, setUnlimitedParticipants, setUseSplitCapacities,
      setWaitlistEnabled,
    }, ev);
  };

  // v24.5: Demo-Events finden immer am NÄCHSTEN Samstag statt.
  const nextSaturdayAt = (hour: number, minute: number): Date => {
    const d = new Date();
    const day = d.getDay(); // 0=So … 6=Sa
    let add = (6 - day + 7) % 7; // Tage bis zum kommenden Samstag
    if (add === 0) add = 7;      // heute Samstag → nächster Samstag
    d.setDate(d.getDate() + add);
    d.setHours(hour, minute, 0, 0);
    return d;
  };
  // Tag(e) vor dem nächsten Samstag — für Anmelde-/Abmeldefristen.
  const beforeNextSaturday = (daysBefore: number, hour: number, minute: number): Date => {
    const d = nextSaturdayAt(hour, minute);
    d.setDate(d.getDate() - daysBefore);
    return d;
  };

  // v11.88: Vier Demo-Vorlagen — vom „Demo"-Button-Modal aufgerufen.
  // Jede Variante füllt das Formular vollständig (inkl. Reset der
  // Felder, die diese Variante NICHT setzt).
  const loadDemoStandard = (): void => {
    resetDemoVariantBaseState();
    const start = nextSaturdayAt(10, 0);
    const end = nextSaturdayAt(12, 0);
    const deadline = beforeNextSaturday(1, 23, 59);
    setTitle('Demo-Meeting Standard');
    setDescription('Beispielhaftes einfaches Meeting ohne Gruppen und ohne Sub-Events.');
    setLocation('Heinrich Campus Düsseldorf, 6. Etage');
    setStartDate(fmtDatetime(start));
    setEndDate(fmtDatetime(end));
    setRegistrationDeadline(fmtDate(deadline));
    setLastDeregisterDate(fmtDate(deadline));
    setMaxParticipants('50');
    setUseSplitCapacities(false);
    setWaitlistEnabled(true);
    setAskSalutation(false);
    const tDemo = Date.now();
    setCustomFields([
      { id: `cf-${tDemo}`, label: 'Essenspräferenz', type: 'select', required: true,
        options: ['Vegetarisch', 'Vegan', 'Keine Einschränkungen'], visible: true },
    ]);
    setCurrentStep(0);
  };

  const loadDemoGroups = (): void => {
    resetDemoVariantBaseState();
    const start = nextSaturdayAt(9, 0);
    const end = nextSaturdayAt(17, 0);
    const deadline = beforeNextSaturday(2, 23, 59);
    setTitle('Demo-Workshop mit Gruppen');
    setDescription('Workshop mit zwei Teilnehmer-Gruppen (Vormittag/Nachmittag) und gemeinsamer Warteliste.');
    setLocation('Deloitte Office Köln');
    setStartDate(fmtDatetime(start));
    setEndDate(fmtDatetime(end));
    setRegistrationDeadline(fmtDate(deadline));
    setLastDeregisterDate(fmtDate(deadline));
    setMaxParticipants('50');
    setUseSplitCapacities(true);
    setSplitLabelA('Vormittag');
    setSplitLabelB('Nachmittag');
    setDurchstarterCapacity('25');
    setFunstarterCapacity('25');
    setSplitSharedWaitlist(true);
    setWaitlistEnabled(true);
    setAskSalutation(false);
    setCurrentStep(0);
  };


  return {
    SUGGESTED_FIELDS_CATALOG, addCustomField, addSelectedSuggestedFields, addStartblock, addSubEventCustomField, allowAttendeeUpload,
    applyEventTemplate, askSalutation, askTeamName, attendeeUploadHint, attendeeUploadLabel, beforeNextSaturday,
    bilingualFields, confirmDialogEnabled, confirmDialogMode, confirmDialogText, copyParentFieldsToSubEvent, dragOverSectionId,
    dragSectionId, durchstarterRequiresProof, durchstarterStartblock, error, fmtDate, fmtDatetime,
    funstarterStartblock, imageUploadError, isB2runTemplate, isSubmitting, isVisOpen, loadDemoGroups,
    loadDemoStandard, locationOptions, mainEventLabel, mainEventLabelMode, nextSaturdayAt, openSuggestedModal,
    previewSections, progress, progressLabel, registrationLanguage, removeCustomField, removeStartblock,
    removeSubEventCustomField, renderShowIfConfig, resetDemoVariantBaseState, setAllowAttendeeUpload, setAskSalutation, setAskTeamName,
    setAttendeeUploadHint, setAttendeeUploadLabel, setBilingualFields, setConfirmDialogEnabled, setConfirmDialogMode, setConfirmDialogText,
    setDragOverSectionId, setDragSectionId, setDurchstarterStartblock, setError, setFunstarterStartblock, setImageUploadError,
    setIsSubmitting, setMainEventLabel, setMainEventLabelMode, setPreviewSections, setProgress, setProgressLabel,
    setRegistrationLanguage, setShowB2runSuggested, setShowPreview, setShowRegisterPreview, setShowSuggestedModal, setSubmitted,
    setSuggestedSelection, setTeamJoinRequiresApproval, setTeamMembersCannotCreate, setTeamOpenSlotsVisible, setTeamPartialAllowed, setTeamRegistrationEnabled,
    setTeamSize, setTeamTermPlural, setTeamTermSingular, setTriedNext, setUseSplitCapacities, setVisCopyModalOpen,
    showB2runSuggested, showPreview, showRegisterPreview, showSuggestedModal, submitted, suggestedSelection,
    teamJoinRequiresApproval, teamMembersCannotCreate, teamOpenSlotsVisible, teamPartialAllowed, teamRegistrationEnabled, teamSize,
    teamTermPlural, teamTermSingular, triedNext, updateCustomField, updateSubEventCustomField, useSplitCapacities,
    visCopyModalOpen, visCopyPendingRef, visHeader, visSnapshotRef,
  };
}
