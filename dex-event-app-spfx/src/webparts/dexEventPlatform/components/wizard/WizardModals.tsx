/* WizardModals — aus EventCreationPage.tsx ausgelagert (Zeilen 6907-8577 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; die Komponente
 * gibt ein Fragment zurueck, damit die Geschwister-Reihenfolge im Elternbaum
 * exakt bleibt.
 * v31.2: Alle Dialoge auf den einheitlichen Modal-Kopf/-Fuß (title/subtitle/
 * icon/footer) und die dex-ui-Klassen umgestellt — Ergebnis bzw. Frage zuerst,
 * Details darunter, Aktionen im Fuß. Handler, Bedingungen und State-Bindungen
 * sind unverändert; nur Darstellung, Reihenfolge und Wortwahl. */
import * as React from 'react';
import { AlertCircle, Calendar, Check, ChevronDown, Copy, Download, FileText, Info, Plus, Send, Star, Users, X } from '../Icons';
// v31.2: Gemeinsame UI-Klassen (Zeilen, Chips, Hinweiskästen) — siehe dexUi.ts
// und docs/ui-leitfaden.md.
import { cx } from '../dexUi';
import { EmailOverrideEntry } from '../wizard/emailOverrideEntry';
import { buildOutlookLocation } from '../../utils/eventFormat';
import DatePicker from 'react-datepicker';
import { HtmlEditorModal } from '../HtmlEditorModal';
import { shortSubEventTitle } from '../../utils/subEventTitle';
import { DESCRIPTION_TEMPLATES } from '../../data/descriptionTemplates';
import { formatOrganizerList } from '../../context/EventContext';
import { RegisterPreviewModal } from '../RegisterPreviewModal';
import BulkUserImportModal from '../BulkUserImportModal';
import { SummaryData, exportSummaryAsDoc, exportSummaryAsPdf } from '../../services/EventSummaryExport';
import Modal from '../Modal';
import { OutlookConfirmItem, SubEventDraft, SuggestedEntry } from '../wizard/wizardTypes';
import { Icon } from '@fluentui/react/lib/Icon';
import { InfoTooltip } from '../InfoTooltip';
import { AgendaItem } from '../../types';
import { CustomFieldInput } from '../wizard/customFieldInput';
import { outlookDefaultBodyTemplate } from '../../utils/outlookDefaultBody';
import { buildProgramHtml } from '../../utils/programPlaceholder';

export interface WizardModalsProps {
  activeCommTabIdx: number;
  activeFrom: string;
  addrCity: string;
  addrHouseNo: string;
  addrStreet: string;
  addrZip: string;
  addSelectedSuggestedFields: () => void;
  agenda: AgendaItem[];
  applySubTransfer: () => void;
  askSalutation: boolean;
  attemptSubmit: () => void;
  audience: string;
  berlinLocalToUtcIso: (localStr: string) => string;
  bilingualFields: boolean;
  buildDraftPayload: () => Record<string, unknown>;
  bulkOrganizerOpen: boolean;
  bulkQrScannerOpen: boolean;
  bulkTestTeamOpen: boolean;
  cancelOutlookSave: () => void;
  childTermPlural: string;
  childTermSingular: string;
  closeVisCopy: (apply: boolean) => void;
  confirmOutlookSave: () => void;
  contactEmail: string;
  customFields: CustomFieldInput[];
  DEMO_VARIANTS: Record<"standard" | "groups" | "subevent" | "subeventTeam", () => void>;
  description: string;
  disableEmails: boolean;
  disableOutlook: boolean;
  documents: { name: string; file?: File; url: string; size: number; }[];
  DRAFT_KEY: string;
  dragOverSectionId: string;
  dragSectionId: string;
  durchstarterCapacity: string;
  emailLanguage: string;
  emailLogoPreview: string;
  emailTemplateOverrides: Record<string, EmailOverrideEntry>;
  emailTemplates: { id: number; templateType: string; language: string; subject: string; heading: string; headingColor: string; bodyHtml: string; }[];
  endDate: string;
  eventImageUrl: string;
  excludedUsers: string[];
  filterMode: "AND" | "OR";
  funstarterCapacity: string;
  headerImageLayout: { width: number; paddingV: number; paddingH: number; };
  htmlEditorMode: "outlook" | "email" | "description";
  htmlEditorOpen: boolean;
  htmlEditorTemplateType: string;
  imagePreview: string;
  isDe: boolean;
  isEditMode: boolean;
  isFictive: boolean;
  isMobile: boolean;
  isoToLocal: (iso: string) => string;
  lastDeregisterDate: string;
  location: string;
  locationFilter: string;
  maxParticipants: string;
  newSectionError: string;
  newSectionModalOpen: boolean;
  newSectionName: string;
  organizer: string;
  organizerEmails: string[];
  outlookBody: string;
  outlookConfirmChecks: Record<string, boolean>;
  outlookConfirmItems: OutlookConfirmItem[];
  outlookConfirmOpen: boolean;
  outlookEndOverride: string;
  outlookHeading: string;
  outlookLocationOverride: string;
  outlookLogoPreview: string;
  outlookStartOverride: string;
  outlookSubheading: string;
  outlookSubject: string;
  pendingSections: string[];
  pendingSuccessDispatch: { title: string; eventId: string; type: "create" | "update"; };
  pendingSuccessDispatchRef: React.MutableRefObject<{ title: string; eventId: string; type: 'create' | 'update'; }>;
  previewSections: { id: string; label: string; }[];
  qrScannerEmails: string[];
  qrScannerNames: string[];
  quiz: { id: string; question: string; options: string[]; correctIndices: number[]; imageBase64?: string; section?: string; }[];
  registrationDeadline: string;
  registrationLanguage: "" | "de" | "en";
  renderPreviewSection: (sectionId: string) => React.ReactElement | null;
  requireSubEventSelection: boolean;
  // v30.67: Der Prüfen-Dialog liest die Kommunikation des HAUPTEVENTS über
  // den Resolver, nicht aus den Spiegel-States (die auf einem Sub-Reiter
  // den Sub-Wert tragen).
  resolveTopLevelCommState: () => { emailLanguage: string; disableEmails: boolean; disableRegistrationEmail: boolean; disableCancellationEmail: boolean; autoDeregisterOnDecline: boolean; inactiveHandling?: 'notify' | 'autoderegister'; disableOutlook: boolean; };
  scDescription: string;
  scopeSub: SubEventDraft;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<{ email: string; displayName: string; location: string; jobTitle: string; }[]>;
  setBulkOrganizerOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setBulkQrScannerOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setBulkTestTeamOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setDragOverSectionId: React.Dispatch<React.SetStateAction<string>>;
  setDragSectionId: React.Dispatch<React.SetStateAction<string>>;
  setEmailTemplateOverrides: React.Dispatch<React.SetStateAction<Record<string, EmailOverrideEntry>>>;
  setHeaderImageLayout: React.Dispatch<React.SetStateAction<{ width: number; paddingV: number; paddingH: number; }>>;
  setHtmlEditorOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setNewSectionError: React.Dispatch<React.SetStateAction<string>>;
  setNewSectionModalOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setNewSectionName: React.Dispatch<React.SetStateAction<string>>;
  setOrganizer: React.Dispatch<React.SetStateAction<string>>;
  setOrganizerEmails: React.Dispatch<React.SetStateAction<string[]>>;
  setOutlookBody: React.Dispatch<React.SetStateAction<string>>;
  setOutlookConfirmChecks: React.Dispatch<React.SetStateAction<Record<string, boolean>>>;
  setOutlookEndOverride: React.Dispatch<React.SetStateAction<string>>;
  setOutlookHeading: React.Dispatch<React.SetStateAction<string>>;
  setOutlookLocationOverride: React.Dispatch<React.SetStateAction<string>>;
  setOutlookStartOverride: React.Dispatch<React.SetStateAction<string>>;
  setOutlookSubheading: React.Dispatch<React.SetStateAction<string>>;
  setOutlookSubject: React.Dispatch<React.SetStateAction<string>>;
  setPendingSections: React.Dispatch<React.SetStateAction<string[]>>;
  setPendingSuccessDispatch: React.Dispatch<React.SetStateAction<{ title: string; eventId: string; type: "create" | "update"; }>>;
  setPreviewSections: React.Dispatch<React.SetStateAction<{ id: string; label: string; }[]>>;
  setQrScannerEmails: React.Dispatch<React.SetStateAction<string[]>>;
  setQrScannerNames: React.Dispatch<React.SetStateAction<string[]>>;
  setScDescription: (v: string) => void;
  setShowB2runSuggested: React.Dispatch<React.SetStateAction<boolean>>;
  setShowConfigCheck: React.Dispatch<React.SetStateAction<boolean>>;
  setShowDemoVariantModal: React.Dispatch<React.SetStateAction<boolean>>;
  setShowPreview: React.Dispatch<React.SetStateAction<boolean>>;
  setShowRegisterPreview: React.Dispatch<React.SetStateAction<boolean>>;
  setShowSuggestedModal: React.Dispatch<React.SetStateAction<boolean>>;
  setShowSummaryModal: React.Dispatch<React.SetStateAction<boolean>>;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setSubTransfer: React.Dispatch<React.SetStateAction<{ fromIdx: number; groups: string[]; targets: number[]; }>>;
  setSuggestedSelection: React.Dispatch<React.SetStateAction<Record<string, boolean>>>;
  setTestTeamEmails: React.Dispatch<React.SetStateAction<string[]>>;
  setTestTeamNames: React.Dispatch<React.SetStateAction<string[]>>;
  setUnsavedConfirmOpen: React.Dispatch<React.SetStateAction<{ resolve: (_ok: boolean) => void; }>>;
  showB2runSuggested: boolean;
  showConfigCheck: boolean;
  showDemoVariantModal: boolean;
  showPreview: boolean;
  showRegisterPreview: boolean;
  showSuggestedModal: boolean;
  showSummaryModal: boolean;
  splitLabelA: string;
  splitLabelB: string;
  splitSharedWaitlist: boolean;
  startDate: string;
  SUB_TRANSFER_GROUPS: { key: string; de: string; en: string; fields: string[]; }[];
  subEvents: SubEventDraft[];
  subEventsOnlyMode: boolean;
  subGroupDiffCount: (srcIdx: number, fields: string[]) => number;
  subTransfer: { fromIdx: number; groups: string[]; targets: number[]; };
  SUGGESTED_FIELDS_CATALOG: SuggestedEntry[];
  suggestedSelection: Record<string, boolean>;
  t: (key: string) => string;
  teamRegistrationEnabled: boolean;
  teamSize: number;
  testTeamEmails: string[];
  testTeamNames: string[];
  title: string;
  transferTimes: { id: string; location: string; meetingPoint: string; address: string; date: string; departureTime: string; arrivalTime: string; description: string; }[];
  unlimitedParticipants: boolean;
  unsavedConfirmOpen: { resolve: (_ok: boolean) => void; };
  useSplitCapacities: boolean;
  visCopyModalOpen: boolean;
  waitlistEnabled: boolean;
  allowAttendeeUpload: boolean;
  askTeamName: boolean;
  attendeeUploadHint: string;
  attendeeUploadLabel: string;
  contactInfo: string;
  contactName: string;
  notifyOrgCancelMode: "never" | "always" | "afterDeadline";
  notifyOrgRegisterFromDate: string;
  notifyOrgRegisterMode: "never" | "always" | "fromDate";
  quizClusterSize: number;
  splitDescA: string;
  splitDescB: string;
  splitDisplayOrderReversed: boolean;
  splitHelpText: string;
  splitSectionTitle: string;
  teamJoinRequiresApproval: boolean;
  teamOpenSlotsVisible: boolean;
  teamPartialAllowed: boolean;
}

export const WizardModals: React.FC<WizardModalsProps> = (p) => {
  const { activeCommTabIdx, activeFrom, addrCity, addrHouseNo, addrStreet, addrZip, addSelectedSuggestedFields, agenda, applySubTransfer, askSalutation, attemptSubmit, audience, berlinLocalToUtcIso, bilingualFields, buildDraftPayload, bulkOrganizerOpen, bulkQrScannerOpen, bulkTestTeamOpen, cancelOutlookSave, childTermPlural, childTermSingular, closeVisCopy, confirmOutlookSave, contactEmail, customFields, DEMO_VARIANTS, description, disableEmails, disableOutlook, documents, DRAFT_KEY, dragOverSectionId, dragSectionId, durchstarterCapacity, emailLanguage, emailLogoPreview, emailTemplateOverrides, emailTemplates, endDate, eventImageUrl, excludedUsers, filterMode, funstarterCapacity, headerImageLayout, htmlEditorMode, htmlEditorOpen, htmlEditorTemplateType, imagePreview, isDe, isEditMode, isFictive, isMobile, isoToLocal, lastDeregisterDate, location, locationFilter, maxParticipants, newSectionError, newSectionModalOpen, newSectionName, organizer, organizerEmails, outlookBody, outlookConfirmChecks, outlookConfirmItems, outlookConfirmOpen, outlookEndOverride, outlookHeading, outlookLocationOverride, outlookLogoPreview, outlookStartOverride, outlookSubheading, outlookSubject, pendingSections, pendingSuccessDispatch, pendingSuccessDispatchRef, previewSections, qrScannerEmails, qrScannerNames, quiz, registrationDeadline, registrationLanguage, renderPreviewSection, requireSubEventSelection, resolveTopLevelCommState, scDescription, scopeSub, searchUsers, setBulkOrganizerOpen, setBulkQrScannerOpen, setBulkTestTeamOpen, setDragOverSectionId, setDragSectionId, setEmailTemplateOverrides, setHeaderImageLayout, setHtmlEditorOpen, setNewSectionError, setNewSectionModalOpen, setNewSectionName, setOrganizer, setOrganizerEmails, setOutlookBody, setOutlookConfirmChecks, setOutlookEndOverride, setOutlookHeading, setOutlookLocationOverride, setOutlookStartOverride, setOutlookSubheading, setOutlookSubject, setPendingSections, setPendingSuccessDispatch, setPreviewSections, setQrScannerEmails, setQrScannerNames, setScDescription, setShowB2runSuggested, setShowConfigCheck, setShowDemoVariantModal, setShowPreview, setShowRegisterPreview, setShowSuggestedModal, setShowSummaryModal, setSubEvents, setSubTransfer, setSuggestedSelection, setTestTeamEmails, setTestTeamNames, setUnsavedConfirmOpen, showB2runSuggested, showConfigCheck, showDemoVariantModal, showPreview, showRegisterPreview, showSuggestedModal, showSummaryModal, splitLabelA, splitLabelB, splitSharedWaitlist, startDate, SUB_TRANSFER_GROUPS, subEvents, subEventsOnlyMode, subGroupDiffCount, subTransfer, SUGGESTED_FIELDS_CATALOG, suggestedSelection, t, teamRegistrationEnabled, teamSize, testTeamEmails, testTeamNames, title, transferTimes, unlimitedParticipants, unsavedConfirmOpen, useSplitCapacities, visCopyModalOpen, waitlistEnabled, allowAttendeeUpload, askTeamName, attendeeUploadHint, attendeeUploadLabel, contactInfo, contactName, notifyOrgCancelMode, notifyOrgRegisterFromDate, notifyOrgRegisterMode, quizClusterSize, splitDescA, splitDescB, splitDisplayOrderReversed, splitHelpText, splitSectionTitle, teamJoinRequiresApproval, teamOpenSlotsVisible, teamPartialAllowed } = p;
  // v31.2: Eine Prüf-und-Anlege-Logik für Enter-Taste UND Knopf im Dialog
  // „Neuer Bereich" — vorher stand derselbe Block zweimal. Kein Hook, nur
  // eine Funktion über den destrukturierten Props.
  const submitNewSection = (): void => {
    const name = newSectionName.trim();
    if (!name) { setNewSectionError(isDe ? 'Bitte einen Namen eingeben.' : 'Please enter a name.'); return; }
    const existing = new Set<string>();
    for (const q of quiz) if (q.section) existing.add(q.section);
    for (const s of pendingSections) existing.add(s);
    if (existing.has(name)) { setNewSectionError(isDe ? 'Ein Bereich mit diesem Namen existiert bereits.' : 'A section with this name already exists.'); return; }
    setPendingSections([...pendingSections, name]);
    setNewSectionModalOpen(false);
  };
  return (
    <>
      {/* ===== Vollbild-Vorschau Modal ===== */}
      {showPreview && (
        <div className="preview-modal" role="dialog" aria-modal="true" aria-label={isDe ? 'Vorschau der Anmeldeseite' : 'Registration page preview'} style={{
          position: 'fixed', inset: 0, background: 'var(--dex-gray-50, #fafafa)', zIndex: 1000,
          display: 'flex', flexDirection: 'column',
        }}>
          <div className="preview-modal-inner" style={{
            width: '100%', maxWidth: '100%',
            height: '100%', overflow: 'auto', padding: 0,
          }}>
            <div style={{
              padding: '14px 24px', borderBottom: '1px solid var(--dex-gray-200)',
              display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12,
              position: 'sticky', top: 0, background: '#fff', zIndex: 1,
            }}>
              <div style={{ minWidth: 0 }}>
                <h3 className="dex-ui-modal-title">{isDe ? 'Vorschau der Anmeldeseite' : 'Registration page preview'}</h3>
                <p className="dex-ui-modal-subtitle">
                  {isDe ? 'So sehen Teilnehmer dein Formular. Zieh einen Abschnitt am Griff, um die Reihenfolge zu ändern.' : 'This is what attendees see. Drag a section by its handle to change the order.'}
                </p>
              </div>
              <button type="button" className="dex-ui-iconbtn" aria-label={isDe ? 'Schließen' : 'Close'} title={isDe ? 'Schließen' : 'Close'} onClick={() => setShowPreview(false)}>
                <X size={20} />
              </button>
            </div>

            <div style={{ padding: 24, display: 'flex', flexDirection: 'column', gap: 16 }}>
              {previewSections.map(section => (
                <div
                  key={section.id}
                  draggable
                  onDragStart={() => setDragSectionId(section.id)}
                  onDragOver={(e) => { e.preventDefault(); setDragOverSectionId(section.id); }}
                  onDragLeave={() => { if (dragOverSectionId === section.id) setDragOverSectionId(null); }}
                  onDrop={() => {
                    if (dragSectionId && dragSectionId !== section.id) {
                      const fromIdx = previewSections.findIndex(s => s.id === dragSectionId);
                      const toIdx = previewSections.findIndex(s => s.id === section.id);
                      if (fromIdx >= 0 && toIdx >= 0) {
                        const updated = [...previewSections];
                        const [moved] = updated.splice(fromIdx, 1);
                        updated.splice(toIdx, 0, moved);
                        setPreviewSections(updated);
                      }
                    }
                    setDragSectionId(null);
                    setDragOverSectionId(null);
                  }}
                  onDragEnd={() => { setDragSectionId(null); setDragOverSectionId(null); }}
                  style={{
                    opacity: dragSectionId === section.id ? 0.4 : 1,
                    // v31.2: Einwurf-Marke als Schatten statt Rahmen — ein
                    // borderTop ließ den Abschnitt beim Überfahren springen.
                    boxShadow: dragOverSectionId === section.id ? 'inset 0 3px 0 var(--dex-green, #86bc25)' : undefined,
                    borderRadius: 12,
                    cursor: 'grab',
                    position: 'relative',
                    transition: 'opacity 0.15s ease, box-shadow 0.15s ease',
                  }}
                >
                  <div
                    className="dex-ui-drag-handle"
                    title={isDe ? 'Zum Verschieben ziehen' : 'Drag to move'}
                    style={{ position: 'absolute', top: 4, right: 8, fontSize: '0.7rem', fontWeight: 600, userSelect: 'none', zIndex: 1, gap: 4 }}
                  >
                    ⠿ {isDe ? 'verschieben' : 'move'}
                  </div>
                  {renderPreviewSection(section.id)}
                </div>
              ))}
            </div>

            <div style={{
              padding: '14px 24px', borderTop: '1px solid var(--dex-gray-200)',
              display: 'flex', gap: 12, justifyContent: 'flex-end', flexWrap: 'wrap',
              position: 'sticky', bottom: 0, background: '#fff',
            }}>
              <button type="button" className="btn btn-secondary" onClick={() => setShowPreview(false)}>
                {isDe ? 'Zurück zum Formular' : 'Back to the form'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                disabled={!title}
                title={!title ? (isDe ? 'Ohne Event-Titel kann nicht gespeichert werden' : 'A title is required before saving') : undefined}
                onClick={() => { setShowPreview(false); attemptSubmit(); }}
              >
                <Send size={16} /> {isEditMode ? t('create.save') : t('create.submit')}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* HTML-Editor-Modal mit Live-Preview (Outlook-Termin, E-Mail-Template oder Beschreibung).
          v9.39: Mode 'description' für die Event-Beschreibung — wird auf der Anmelde-Seite
          1:1 als HTML gerendert, deshalb hier auch ein Bearbeiten/Vorschau-Modal wie bei den
          Mail-Templates. */}
      {(() => {
        if (!htmlEditorOpen) return null;
        const isOutlook = htmlEditorMode === 'outlook';
        const isDescription = htmlEditorMode === 'description';
        const tType = htmlEditorTemplateType;
        const defaultTpl = (!isOutlook && !isDescription) ? emailTemplates.find(tp => tp.templateType === tType && tp.language === emailLanguage) : undefined;
        const override = (!isOutlook && !isDescription) ? emailTemplateOverrides[tType] : undefined;
        const currentSubject = override?.subject || defaultTpl?.subject || '';
        const currentHeading = override?.heading || defaultTpl?.heading || '';
        // v15.19: Subheading-Override pro Event. Falls override.subheading
        // explizit gesetzt ist (auch leerer String), nutze diesen Wert.
        const currentSubheading = override?.subheading !== undefined ? override.subheading : '';
        // v28.89: Im Beschreibungs-Modus folgt der Editor der gewählten Ebene
        // (scDescription/setScDescription) — er wird ausschließlich aus
        // Schritt 1 geöffnet, wo der Scope-Reiter darüber steht.
        const currentBody = isOutlook
          ? outlookBody
          : isDescription
            ? scDescription
            : (override?.bodyHtml || defaultTpl?.bodyHtml || '');
        // v18.19: Überschrift-Farbe + -Größe (Override > Template-Default).
        const currentHeadingColor = (override?.headingColor) || (defaultTpl?.headingColor) || '#86bc25';
        const currentHeadingFontSize = override?.headingFontSize || '26px';
        // v18.22: Überschrift fett/kursiv + Unter-Überschrift-Formatierung.
        const currentHeadingBold = override?.headingBold;
        const currentHeadingItalic = override?.headingItalic;
        const currentSubheadingColor = override?.subheadingColor || '#000000';
        const currentSubheadingFontSize = override?.subheadingFontSize || '20px';
        const currentSubheadingBold = override?.subheadingBold;
        const currentSubheadingItalic = override?.subheadingItalic;
        // v18.22: zentraler Patch-Helper — merged ein Teil-Update in den
        // Override des aktuellen TemplateTypes und BEWAHRT alle übrigen Felder
        // (vorher droppte z.B. ein Heading-Text-Edit die zuvor gesetzte Farbe).
        const patchOverride = (patch: Partial<EmailOverrideEntry>): void => {
          setEmailTemplateOverrides(prev => {
            const cur = prev[tType];
            return {
              ...prev,
              [tType]: {
                subject: cur?.subject ?? currentSubject,
                heading: cur?.heading ?? currentHeading,
                subheading: cur?.subheading !== undefined ? cur.subheading : currentSubheading,
                bodyHtml: cur?.bodyHtml ?? currentBody,
                ...(cur ? {
                  headingColor: cur.headingColor,
                  headingFontSize: cur.headingFontSize,
                  headingBold: cur.headingBold,
                  headingItalic: cur.headingItalic,
                  subheadingColor: cur.subheadingColor,
                  subheadingFontSize: cur.subheadingFontSize,
                  subheadingBold: cur.subheadingBold,
                  subheadingItalic: cur.subheadingItalic,
                } : {}),
                ...patch,
              },
            };
          });
        };
        // v18.42: read-only Termin/Ort-Labels für den Outlook-Editor — je nach
        // aktivem Tab (Hauptevent oder Sub-Event).
        const olActiveSub = activeCommTabIdx > 0 ? subEvents[activeCommTabIdx - 1] : undefined;
        const olStart = olActiveSub ? olActiveSub.startDate : startDate;
        const olEnd = olActiveSub ? olActiveSub.endDate : endDate;
        const olFmt = (d?: string): string => {
          if (!d) return '';
          try { return new Date(d).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); } catch { return ''; }
        };
        // v18.44: Auto-Ort (= „würde übernommen") als Platzhalter.
        const outlookLocationAuto = olActiveSub
          ? (buildOutlookLocation(olActiveSub.location, olActiveSub.locationAddress) || olActiveSub.location || '')
          : (buildOutlookLocation(location, { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity }));
        // v18.44: aktuelle Override-Werte des aktiven Tabs (leer = übernommen).
        const olLocationOverrideVal = olActiveSub ? (olActiveSub.outlookLocation || '') : outlookLocationOverride;
        const olStartOverrideVal = olActiveSub ? (olActiveSub.outlookStart || '') : outlookStartOverride;
        const olEndOverrideVal = olActiveSub ? (olActiveSub.outlookEnd || '') : outlookEndOverride;
        const setOlLocation = (v: string): void => {
          if (olActiveSub) { const fi = activeCommTabIdx - 1; setSubEvents(prev => prev.map((s, i) => i === fi ? { ...s, outlookLocation: v } : s)); }
          else setOutlookLocationOverride(v);
        };
        const setOlStart = (iso: string): void => {
          if (olActiveSub) { const fi = activeCommTabIdx - 1; setSubEvents(prev => prev.map((s, i) => i === fi ? { ...s, outlookStart: iso } : s)); }
          else setOutlookStartOverride(iso);
        };
        const setOlEnd = (iso: string): void => {
          if (olActiveSub) { const fi = activeCommTabIdx - 1; setSubEvents(prev => prev.map((s, i) => i === fi ? { ...s, outlookEnd: iso } : s)); }
          else setOutlookEndOverride(iso);
        };
        const pad2 = (n: number): string => String(n).padStart(2, '0');
        const olIsoToDate = (iso?: string): Date | null => {
          if (!iso) return null;
          const loc = isoToLocal(iso); if (!loc) return null;
          const [dp, tp] = loc.split('T'); const [y, mo, da] = dp.split('-').map(n => parseInt(n, 10)); const [h, mi] = (tp || '00:00').split(':').map(n => parseInt(n, 10));
          return new Date(y, mo - 1, da, h, mi, 0, 0);
        };
        const olDateToIso = (d: Date | null): string => {
          if (!d) return '';
          return berlinLocalToUtcIso(`${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}T${pad2(d.getHours())}:${pad2(d.getMinutes())}`);
        };
        const dpCommon = {
          showTimeSelect: true, timeFormat: 'HH:mm', timeIntervals: 15, timeCaption: isDe ? 'Uhrzeit' : 'Time',
          dateFormat: 'dd.MM.yyyy, HH:mm', locale: 'de', className: 'form-input',
          wrapperClassName: 'dex-datepicker-wrapper', calendarClassName: 'dex-datepicker-calendar',
          popperPlacement: 'bottom-start' as const, isClearable: true, autoComplete: 'off',
        };
        // v31.2: Der Platzhalter sagt, welcher Wert ohne Eingabe gilt
        // („… (aus dem Event)") — leer heißt hier nicht „kein Termin", sondern
        // „wie in Schritt 1".
        const inherited = isDe ? ' (aus dem Event)' : ' (from the event)';
        const outlookDateEditor = (
          <div className="form-grid-2col dex-ui-grid-2" style={{ gap: 8 }}>
            <div className="dex-ui-field" style={{ marginBottom: 0 }}>
              <label className="dex-ui-label" style={{ fontSize: '0.78rem', marginBottom: 4 }}>{isDe ? 'Beginn im Termin' : 'Start in the invite'}</label>
              <DatePicker {...dpCommon} selected={olIsoToDate(olStartOverrideVal)} onChange={(d: Date | null) => setOlStart(olDateToIso(d))} placeholderText={olStart ? olFmt(olStart) + inherited : (isDe ? 'Beginn' : 'Start')} />
            </div>
            <div className="dex-ui-field" style={{ marginBottom: 0 }}>
              <label className="dex-ui-label" style={{ fontSize: '0.78rem', marginBottom: 4 }}>{isDe ? 'Ende im Termin' : 'End in the invite'}</label>
              <DatePicker {...dpCommon} selected={olIsoToDate(olEndOverrideVal)} onChange={(d: Date | null) => setOlEnd(olDateToIso(d))} placeholderText={olEnd ? olFmt(olEnd) + inherited : (isDe ? 'Ende' : 'End')} />
            </div>
          </div>
        );
        // v18.46: Standard-Body-Vorlage (mit Platzhaltern) für „Standardtext laden"
        // im Outlook-Editor — Sprache folgt der aktiven Mail-Sprache.
        // v30.94: dieselbe Vorlage wie Speichern und Vorschau-Karte (utils/outlookDefaultBody).
        const outlookDefaultBody = outlookDefaultBodyTemplate(emailLanguage);
        // v19.2: Einladender Beispieltext für die Beschreibung — über den
        // „Standardtext laden"-Button im Beschreibungs-Editor übernehmbar (statt
        // wie früher als Inline-Box im Wizard).
        const descriptionExampleHtml = isDe
          ? 'Liebe Kolleginnen und Kollegen,<br><br>wir freuen uns sehr, euch herzlich einzuladen! Es erwartet euch ein abwechslungsreiches Programm mit viel Raum für Austausch und Begegnung.<br><br>Wir freuen uns auf einen schönen gemeinsamen Tag mit euch!'
          : 'Dear colleagues,<br><br>we are delighted to invite you! Look forward to a varied programme with plenty of room for exchange and networking.<br><br>We look forward to seeing you there!';
        return (
          <HtmlEditorModal
            open={htmlEditorOpen}
            onClose={() => setHtmlEditorOpen(false)}
            defaultBodyHtml={isOutlook ? outlookDefaultBody : (isDescription ? descriptionExampleHtml : undefined)}
            title={isOutlook
              ? (isDe ? 'Outlook-Termin: Text bearbeiten' : 'Outlook invite: edit text')
              : isDescription
                ? (scopeSub
                  ? (isDe
                    ? `Beschreibung: ${shortSubEventTitle(scopeSub.title, title) || (childTermSingular || 'Sub-Event')}`
                    : `Description: ${shortSubEventTitle(scopeSub.title, title) || (childTermSingular || 'sub-event')}`)
                  : (isDe ? 'Event-Beschreibung bearbeiten' : 'Edit event description'))
                : (isDe ? `Mail-Vorlage: ${tType}` : `Email template: ${tType}`)}
            // v28.7: Die Starthilfe (Tipp-Text + Vorschlags-Chips) lebt jetzt
            // HIER im Editor statt als Dauer-Box im Wizard-Schritt.
            headerExtra={isDescription ? (
              <div className="dex-ui-callout dex-ui-callout--info">
                <span className="dex-ui-callout-icon"><Info size={16} /></span>
                <span>
                  {isDe
                    ? <>Die Beschreibung ist der <strong>einladende Einstieg ganz oben auf der Anmeldeseite</strong> — das Erste, was deine Teilnehmenden lesen. Erzähl, <strong>worum es geht, für wen das Event ist und was man wissen sollte</strong>. <strong>Zeit, Ort, Organizer und Kontaktperson kannst du weglassen</strong> — die zeigt die App darüber als eigene Felder.</>
                    : <>The description is the <strong>inviting intro right at the top of the registration page</strong> — the first thing your attendees read. Tell them <strong>what the event is about, who it&rsquo;s for and what to know</strong>. <strong>You can skip date, location, organizer and contact person</strong> — the app shows those above as their own fields.</>}
                </span>
              </div>
            ) : undefined}
            bodyTemplates={isDescription ? DESCRIPTION_TEMPLATES.map(tpl => ({
              key: tpl.key,
              label: isDe ? tpl.labelDe : tpl.labelEn,
              html: isDe ? tpl.de : tpl.en,
              title: (isDe ? tpl.de : tpl.en).replace(/<[^>]+>/g, '').replace(/&rsquo;/g, '’'),
            })) : undefined}
            bodyTemplatesLabel={isDescription ? (isDe ? 'Vorschläge zum Übernehmen (danach frei anpassbar):' : 'Suggestions to use (fully editable afterwards):') : undefined}
            value={currentBody}
            onChange={(html) => {
              if (isOutlook) {
                setOutlookBody(html);
              } else if (isDescription) {
                setScDescription(html);
              } else {
                // v18.22: patchOverride bewahrt alle übrigen Override-Felder
                // (Farbe/Größe/fett/kursiv von Über-/Unter-Überschrift).
                patchOverride({ bodyHtml: html });
              }
            }}
            previewMode={isDescription ? 'plain' : (isOutlook ? 'outlook' : 'email')}
            emailSubject={(!isOutlook && !isDescription) ? currentSubject : undefined}
            onEmailSubjectChange={(!isOutlook && !isDescription) ? (s) => patchOverride({ subject: s }) : undefined}
            emailHeading={(!isOutlook && !isDescription) ? currentHeading : undefined}
            onEmailHeadingChange={(!isOutlook && !isDescription) ? (h) => patchOverride({ heading: h }) : undefined}
            emailSubheading={(!isOutlook && !isDescription) ? currentSubheading : undefined}
            onEmailSubheadingChange={(!isOutlook && !isDescription) ? (s) => patchOverride({ subheading: s }) : undefined}
            emailHeadingColor={(!isOutlook && !isDescription) ? currentHeadingColor : undefined}
            emailHeadingFontSize={(!isOutlook && !isDescription) ? currentHeadingFontSize : undefined}
            onEmailHeadingColorChange={(!isOutlook && !isDescription) ? (hex) => patchOverride({ headingColor: hex }) : undefined}
            onEmailHeadingFontSizeChange={(!isOutlook && !isDescription) ? (px) => patchOverride({ headingFontSize: px }) : undefined}
            emailHeadingBold={(!isOutlook && !isDescription) ? currentHeadingBold : undefined}
            emailHeadingItalic={(!isOutlook && !isDescription) ? currentHeadingItalic : undefined}
            onEmailHeadingBoldChange={(!isOutlook && !isDescription) ? (b) => patchOverride({ headingBold: b }) : undefined}
            onEmailHeadingItalicChange={(!isOutlook && !isDescription) ? (b) => patchOverride({ headingItalic: b }) : undefined}
            emailSubheadingColor={(!isOutlook && !isDescription) ? currentSubheadingColor : undefined}
            emailSubheadingFontSize={(!isOutlook && !isDescription) ? currentSubheadingFontSize : undefined}
            emailSubheadingBold={(!isOutlook && !isDescription) ? currentSubheadingBold : undefined}
            emailSubheadingItalic={(!isOutlook && !isDescription) ? currentSubheadingItalic : undefined}
            onEmailSubheadingColorChange={(!isOutlook && !isDescription) ? (hex) => patchOverride({ subheadingColor: hex }) : undefined}
            onEmailSubheadingFontSizeChange={(!isOutlook && !isDescription) ? (px) => patchOverride({ subheadingFontSize: px }) : undefined}
            onEmailSubheadingBoldChange={(!isOutlook && !isDescription) ? (b) => patchOverride({ subheadingBold: b }) : undefined}
            onEmailSubheadingItalicChange={(!isOutlook && !isDescription) ? (b) => patchOverride({ subheadingItalic: b }) : undefined}
            imageWidth={!isDescription ? headerImageLayout.width : undefined}
            imagePaddingV={!isDescription ? headerImageLayout.paddingV : undefined}
            imagePaddingH={!isDescription ? headerImageLayout.paddingH : undefined}
            onImageWidthChange={!isDescription ? (w) => setHeaderImageLayout(p => ({ ...p, width: w })) : undefined}
            onImagePaddingVChange={!isDescription ? (v) => setHeaderImageLayout(p => ({ ...p, paddingV: v })) : undefined}
            onImagePaddingHChange={!isDescription ? (h) => setHeaderImageLayout(p => ({ ...p, paddingH: h })) : undefined}
            outlookHeading={isOutlook ? outlookHeading : undefined}
            onOutlookHeadingChange={isOutlook ? setOutlookHeading : undefined}
            outlookSubheading={isOutlook ? outlookSubheading : undefined}
            onOutlookSubheadingChange={isOutlook ? setOutlookSubheading : undefined}
            outlookSubject={isOutlook ? outlookSubject : undefined}
            onOutlookSubjectChange={isOutlook ? setOutlookSubject : undefined}
            outlookDateEditor={isOutlook ? outlookDateEditor : undefined}
            outlookLocationValue={isOutlook ? olLocationOverrideVal : undefined}
            onOutlookLocationChange={isOutlook ? setOlLocation : undefined}
            outlookLocationAuto={isOutlook ? outlookLocationAuto : undefined}
            previewVars={{
              // v17.5: Im Sub-Event-Kommunikations-Tab den Titel des
              // aktiven Sub-Events einsetzen, sonst den Hauptevent-Titel.
              EventTitle: (() => {
                if (activeCommTabIdx > 0) {
                  const sub = subEvents[activeCommTabIdx - 1];
                  return (sub && sub.title && sub.title.trim()) || title || 'Event Title';
                }
                return title || 'Event Title';
              })(),
              Name: 'Max Mustermann',
              // v27.5: normalisierte Organizer-Namen ("Vorname Nachname" + „und").
              Organizer: formatOrganizerList([organizer], emailLanguage) || organizer || 'Organisator',
              ContactEmail: contactEmail.trim() || 'kontakt@deloitte.de',
              AppUrl: 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView',
              WaitlistPosition: '1',
              Address: [addrStreet, addrHouseNo].filter(Boolean).join(' ') + ((addrZip || addrCity) ? ', ' + [addrZip, addrCity].filter(Boolean).join(' ') : ''),
              Location: location || 'Veranstaltungsort',
              StartDate: startDate ? new Date(startDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
              EndDate: endDate ? new Date(endDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
              EventDate: startDate ? new Date(startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
            }}
            // v30.95: {{Programm}} als fertiges HTML in die Vorschau (roh, wie QR_BLOCK).
            previewHtmlVars={{ Programm: buildProgramHtml(activeCommTabIdx > 0 ? ((subEvents[activeCommTabIdx - 1] || {}).agenda || []) : agenda, emailLanguage) }}
            insertableVars={isOutlook ? [
              // v17.16: {{Name}} hier ENTFERNT — der Outlook-Termin geht
              // an alle Teilnehmer gleichzeitig, eine pro-Person-Anrede
              // ist nicht möglich. Vorher konnte der Organizer {{Name}}
              // einfügen, was bei allen Empfängern als unaufgelöster
              // Platzhalter „{{Name}}" stehen blieb.
              { key: '{{EventTitle}}', label: 'Event' },
              { key: '{{Organizer}}', label: 'Organizer' },
              // v27.5: {{ContactEmail}} nur anbieten, wenn eine Ansprechpartner-
              // Mail hinterlegt ist (Schritt „Ansprechpartner").
              ...(contactEmail.trim() ? [{ key: '{{ContactEmail}}', label: 'Kontakt-Mail' }] : []),
              { key: '{{Location}}', label: 'Ort' },
              { key: '{{Address}}', label: 'Adresse' },
              { key: '{{StartDate}}', label: 'Start' },
              { key: '{{EndDate}}', label: 'Ende' },
              { key: '{{AppUrl}}', label: 'App Link' },
              // v30.95: Programm-Tabelle — nur anbieten, wenn es Punkte gibt.
              ...((activeCommTabIdx > 0 ? ((subEvents[activeCommTabIdx - 1] || {}).agenda || []) : agenda).length ? [{ key: '{{Programm}}', label: 'Programm' }] : []),
            ] : [
              { key: '{{Name}}', label: 'Name' },
              { key: '{{EventTitle}}', label: 'Event' },
              { key: '{{Organizer}}', label: 'Organizer' },
              ...(contactEmail.trim() ? [{ key: '{{ContactEmail}}', label: 'Kontakt-Mail' }] : []),
              { key: '{{AppUrl}}', label: 'App Link' },
              { key: '{{WaitlistPosition}}', label: 'Waitlist #' },
              ...((activeCommTabIdx > 0 ? ((subEvents[activeCommTabIdx - 1] || {}).agenda || []) : agenda).length ? [{ key: '{{Programm}}', label: 'Programm' }] : []),
            ]}
            imageBase64={(isOutlook ? outlookLogoPreview : emailLogoPreview) || ''}
          />
        );
      })()}

      {/* Register-Page-Preview-Modal (zeigt, was Teilnehmer sehen würden) */}
      <RegisterPreviewModal
        open={showRegisterPreview}
        onClose={() => setShowRegisterPreview(false)}
        data={{
          title,
          description,
          location,
          locationAddress: { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity },
          startDate,
          endDate,
          imagePreview,
          organizers: organizer.split(';').map(s => s.trim()).filter(Boolean),
          organizerEmails,
          maxParticipants: Number(maxParticipants) || 0,
          unlimitedParticipants,
          customFields: customFields.map(f => ({
            id: f.id,
            label: f.label,
            type: f.type,
            required: f.required,
            visible: f.visible !== false,
            options: f.type === 'select' ? f.options : undefined,
            // v26.74: Vorauswahl an die Live-Preview weiterreichen.
            defaultValue: f.type === 'select' && !f.multi ? f.defaultValue : undefined,
            // v26.75: Vorfilter-Kategorien + Beschriftung an die Preview.
            optionCategories: f.type === 'select' && !f.multi ? f.optionCategories : undefined,
            prefilterLabel: f.type === 'select' && !f.multi ? f.prefilterLabel : undefined,
            // v7.24: helpText, multi und showIf an die Live-Preview weiterreichen,
            // damit die echte RegistrationPage genau das rendert was der
            // Teilnehmer später sieht (i-Tooltip, Multi-Select-Liste,
            // Sichtbarkeitsbedingung).
            helpText: f.helpText,
            helpTextStyle: f.helpTextStyle,
            multi: f.multi,
            showIf: f.showIf,
            // v17.20: EN-Varianten an die Preview weiterreichen — sonst sieht
            // der Organizer in der Vorschau nicht, was englische Teilnehmer
            // bekommen würden.
            confirmLabel: f.confirmLabel,
            labelEn: f.labelEn,
            helpTextEn: f.helpTextEn,
            confirmLabelEn: f.confirmLabelEn,
            optionsEn: f.optionsEn,
            // v29.21 (Audit): Uhrzeit-Option + Übernachtungs-Fenster an die
            // Vorschau — sonst zeigte sie einen reinen Datums-Picker bzw.
            // einen Zeitraum ohne Grenzen, anders als die echte Anmeldeseite.
            withTime: f.withTime,
            rangeStart: f.rangeStart,
            rangeEnd: f.rangeEnd,
            maxNights: f.maxNights,
          })),
          isFictive,
          // v14.10: Sub-Events + Sub-Only-Mode + Bezeichnungs-Term an die
          // Vorschau weiterreichen, damit der Organizer auch die Sub-Event-
          // Auswahl im Anmeldeformular sieht (vorher fehlte sie komplett).
          subEvents: subEvents.map(s => ({
            id: s.id,
            title: s.title,
            location: s.location,
            startDate: s.startDate,
            endDate: s.endDate,
            maxParticipants: s.maxParticipants,
            description: s.description,
            customFields: (s.customFields || []).map(f => ({
              id: f.id,
              label: f.label,
              type: f.type,
              required: f.required,
              visible: f.visible !== false,
              options: f.type === 'select' ? f.options : undefined,
              helpText: f.helpText,
              helpTextStyle: f.helpTextStyle,
              multi: f.multi,
              showIf: f.showIf,
              // v29.21 (Audit): wie beim Hauptevent — Vorauswahl, EN-Varianten
              // und Datums-Optionen fehlten in der Sub-Event-Vorschau.
              defaultValue: f.type === 'select' && !f.multi ? f.defaultValue : undefined,
              confirmLabel: f.confirmLabel,
              labelEn: f.labelEn,
              helpTextEn: f.helpTextEn,
              confirmLabelEn: f.confirmLabelEn,
              optionsEn: f.optionsEn,
              withTime: f.withTime,
              rangeStart: f.rangeStart,
              rangeEnd: f.rangeEnd,
              maxNights: f.maxNights,
            })),
          })),
          subEventsOnlyMode,
          requireSubEventSelection: requireSubEventSelection || subEventsOnlyMode,
          childEventTermSingular: childTermSingular,
          childEventTermPlural: childTermPlural,
          // v17.22: Bilingual-Flag an die Vorschau — sonst rendert die
          // Preview die EN-Varianten nie (useEnVariants prüft event.bilingualFields).
          bilingualFields,
          // v22.36: Geteilte Kapazität an die Vorschau — sonst fehlt die
          // Gruppenauswahl im Vorschau-Formular.
          ...(useSplitCapacities ? {
            durchstarterCapacity: Number(durchstarterCapacity) || 0,
            funstarterCapacity: Number(funstarterCapacity) || 0,
            splitLabelA,
            splitLabelB,
            splitDescA,
            splitDescB,
            splitHelpText,
            splitSectionTitle,
            splitDisplayOrderReversed,
            splitSharedWaitlist,
          } : {}),
          // v29.21 (Audit): Die Vorschau verspricht „1:1 das, was der
          // Teilnehmer bekommt" — ohne diese Props fehlten Anrede-Dropdown,
          // Gruppen-Beschreibungen und die gedrehte Gruppen-Reihenfolge.
          askSalutation,
        }}
      />

      {/* Massenimport-Modale — eine generische Komponente, mehrere Aufruf-Stellen.
          Teams speichern parallele Names[] + Emails[]-Arrays. Die onAdd-Callbacks
          übersetzen jeweils zwischen Modal-Output (Email + DisplayName) und der
          jeweiligen State-Form. Der Audience-/Sichtbarkeits-Massenimport ist nach
          <AudiencePicker> gewandert (self-contained pro Instanz). */}
      <BulkUserImportModal
        open={bulkOrganizerOpen}
        onClose={() => setBulkOrganizerOpen(false)}
        title={isDe ? 'Mehrere Co-Organizer hinzufügen' : 'Add several co-organizers'}
        description={(
          <p style={{ marginTop: 0 }}>
            {isDe
              ? <>Neue Personen werden <strong>hinten</strong> angehängt — der erste Eintrag der Liste bleibt Haupt-Organizer.</>
              : <>New people are appended at the <strong>end</strong> — the first entry in the list stays the main organizer.</>}
          </p>
        )}
        existingEmails={organizerEmails}
        searchUsers={searchUsers}
        onAdd={({ email, displayName }) => {
          // WICHTIG: functional setState für `organizer`-String, sonst sehen
          // schnelle Sequenz-Calls (Massenimport mit 10+ Namen) alle dieselbe
          // closure-stale Version und nur der letzte Name landet, während
          // organizerEmails über `prev => ...` korrekt akkumuliert. Das führte
          // zu out-of-sync orgNames/orgEmails-Arrays mit falscher Namen-Email-
          // Zuordnung im Duplikat-Hinweis.
          setOrganizer(prev => {
            const existingNames = (prev || '').split(';').map(s => s.trim()).filter(Boolean);
            return [...existingNames, displayName].join('; ');
          });
          setOrganizerEmails(prev => [...prev, email]);
        }}
      />
      <BulkUserImportModal
        open={bulkTestTeamOpen}
        onClose={() => setBulkTestTeamOpen(false)}
        title={isDe ? 'Mehrere Test-Team-Mitglieder hinzufügen' : 'Add several test-team members'}
        description={(
          <p style={{ marginTop: 0 }}>
            {isDe
              ? <>Das Test-Team sieht das Event schon <strong>im Entwurf</strong> und kann sich testweise anmelden.</>
              : <>The test team sees the event <strong>while it is a draft</strong> and can register for testing.</>}
          </p>
        )}
        existingEmails={testTeamEmails}
        searchUsers={searchUsers}
        onAdd={({ email, displayName }) => {
          setTestTeamNames(prev => [...prev, displayName]);
          setTestTeamEmails(prev => [...prev, email]);
        }}
      />
      <BulkUserImportModal
        open={bulkQrScannerOpen}
        onClose={() => setBulkQrScannerOpen(false)}
        title={isDe ? 'Mehrere Check-in-Team-Mitglieder hinzufügen' : 'Add several check-in team members'}
        description={(
          <p style={{ marginTop: 0 }}>
            {isDe
              ? <>Diese Personen dürfen am Eventtag den <strong>QR-Scanner und das Check-in-Tool</strong> benutzen — weitere Organizer-Rechte bekommen sie nicht.</>
              : <>These people may use the <strong>QR scanner and the check-in tool</strong> on the event day — they get no further organizer rights.</>}
          </p>
        )}
        existingEmails={qrScannerEmails}
        searchUsers={searchUsers}
        onAdd={({ email, displayName }) => {
          setQrScannerNames(prev => [...prev, displayName]);
          setQrScannerEmails(prev => [...prev, email]);
        }}
      />

      {/* Das Gruppen-Mitglieder-Modal ist nach <AudiencePicker> gewandert
          (die Audience-Chips, die es öffnen, leben jetzt dort). */}

      {/* v11.88: Demo-Auswahl-Modal — Ersatz für den früheren
          direkten „Demo"-Klick. Vier Karten-Optionen: Standard,
          Mit Gruppen, Mit Sub-Event, Mit Sub-Event + Team. Klick auf
          eine Karte schliesst das Modal und füllt das Formular mit
          der jeweiligen Variante. */}

      {/* v22.36: „Prüfen"-Modal — Übersicht aller Einstellungen: was ist
          gesetzt, wo greifen Standards, welche optionalen Punkte sind leer,
          welche Pflichtangaben fehlen. Rein lesend aus dem Wizard-State. */}
      {showConfigCheck && (() => {
        type CheckStatus = 'ok' | 'default' | 'empty' | 'missing';
        interface CheckRow { label: string; value: React.ReactNode; status: CheckStatus }
        const fmtDt = (v: string): string => {
          if (!v) return '';
          const d = new Date(v);
          return Number.isFinite(d.getTime()) ? d.toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : v;
        };
        const plainDesc = (description || '').replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ').trim();
        const orgList = organizer.split(';').map(s => s.trim()).filter(Boolean);
        const locList = locationFilter.split(',').map(s => s.trim()).filter(Boolean);
        const audList = (audience || '').split(',').map(s => s.trim()).filter(Boolean);
        // v30.67: Kommunikation des Hauptevents auflösen — die Step-6-States
        // (emailLanguage, disableEmails, disableOutlook …) sind Spiegel-States
        // und tragen auf einem Sub-Reiter die Werte DIESES Sub-Events. Der
        // Dialog meldete dann unter „Schritt 6 — Kommunikation" die Mail-
        // Einstellungen des Termins als die des Hauptevents (v11.93-Falle).
        const topComm = resolveTopLevelCommState();
        const sections: Array<{ title: string; rows: CheckRow[] }> = [];
        sections.push({
          title: isDe ? 'Schritt 1 — Grundlagen' : 'Step 1 — Basics',
          rows: [
            { label: isDe ? 'Event-Titel' : 'Event title', value: title || '—', status: title ? 'ok' : 'missing' },
            { label: isDe ? 'Zeitraum' : 'Dates', value: startDate ? `${fmtDt(startDate)}${endDate ? ` – ${fmtDt(endDate)}` : (isDe ? ' (kein Ende — Outlook-Termin nicht möglich)' : ' (no end — Outlook invite not possible)')}` : '—', status: !startDate ? 'missing' : (endDate ? 'ok' : 'empty') },
            { label: isDe ? 'Beschreibung' : 'Description', value: plainDesc ? `${plainDesc.slice(0, 80)}${plainDesc.length > 80 ? '…' : ''}` : '—', status: plainDesc ? 'ok' : 'empty' },
            { label: isDe ? 'Event-Bild' : 'Event image', value: imagePreview ? (isDe ? 'hochgeladen' : 'uploaded') : '—', status: imagePreview ? 'ok' : 'empty' },
            { label: 'Status', value: isFictive ? (activeFrom ? (isDe ? `Entwurf — geht automatisch live am ${fmtDt(activeFrom)}` : `Draft — goes live automatically on ${fmtDt(activeFrom)}`) : (isDe ? 'Entwurf (nur Admins, Organizer, Test-Team)' : 'Draft (admins, organizers, test team only)')) : (isDe ? 'Aktiv — für berechtigte Teilnehmer sichtbar' : 'Active — visible to eligible attendees'), status: isFictive ? 'default' : 'ok' },
          ],
        });
        // v31.2: Die Sub-Events gehören zu Schritt 1 und stehen deshalb direkt
        // hinter den Grundlagen — vorher lagen sie hinter Schritt 2.
        sections.push({
          title: isDe ? 'Schritt 1 — Sub-Events' : 'Step 1 — Sub-events',
          rows: subEvents.length === 0
            ? [{ label: 'Sub-Events', value: isDe ? 'keine' : 'none', status: 'default' }]
            : [
                { label: 'Sub-Events', value: `${subEvents.length} (${subEvents.map(s => s.title || '?').join(', ').slice(0, 90)})`, status: 'ok' },
                { label: isDe ? 'Anmelde-Modus' : 'Registration mode', value: subEventsOnlyMode ? (isDe ? 'Nur für Sub-Events (Klammer nicht buchbar)' : 'Sub-events only (bracket not bookable)') : (isDe ? 'Hauptevent + Sub-Events' : 'Main event + sub-events'), status: 'ok' },
              ],
        });
        sections.push({
          title: isDe ? 'Schritt 2 — Organizer & Team' : 'Step 2 — Organizers & Team',
          rows: [
            { label: 'Organizer', value: orgList.length ? `${orgList.length} ${isDe ? 'Person(en)' : 'person(s)'}` : '—', status: orgList.length ? 'ok' : 'missing' },
            { label: isDe ? 'Test-Team' : 'Test team', value: testTeamEmails.length ? `${testTeamEmails.length} ${isDe ? 'Person(en)' : 'person(s)'}` : '—', status: testTeamEmails.length ? 'ok' : 'empty' },
            { label: isDe ? 'Check-In-Team' : 'Check-in team', value: qrScannerEmails.length ? `${qrScannerEmails.length} ${isDe ? 'Person(en)' : 'person(s)'}` : '—', status: qrScannerEmails.length ? 'ok' : 'empty' },
          ],
        });
        sections.push({
          title: isDe ? 'Schritt 3 — Ort & Programm' : 'Step 3 — Location & programme',
          rows: [
            { label: isDe ? 'Veranstaltungsort' : 'Venue', value: location || '—', status: location ? 'ok' : 'empty' },
            { label: isDe ? 'Adresse' : 'Address', value: (addrStreet || addrCity) ? [addrStreet, addrHouseNo, addrZip, addrCity].filter(Boolean).join(' ') : '—', status: (addrStreet || addrCity) ? 'ok' : 'empty' },
            { label: 'Agenda', value: agenda.length ? `${agenda.length} ${isDe ? 'Programmpunkte' : 'items'}` : '—', status: agenda.length ? 'ok' : 'empty' },
            { label: isDe ? 'Transferzeiten' : 'Transfers', value: transferTimes.length ? `${transferTimes.length}` : '—', status: transferTimes.length ? 'ok' : 'empty' },
          ],
        });
        sections.push({
          title: isDe ? 'Schritt 4 — Kapazität & Sichtbarkeit' : 'Step 4 — Capacity & visibility',
          rows: [
            useSplitCapacities
              ? { label: isDe ? 'Plätze (geteilte Kapazität)' : 'Seats (split capacity)', value: `${splitLabelA || 'Gruppe A'}: ${durchstarterCapacity || 0} · ${splitLabelB || 'Gruppe B'}: ${funstarterCapacity || 0}${splitSharedWaitlist ? (isDe ? ' · gemeinsame Warteliste' : ' · shared waitlist') : ''}`, status: 'ok' }
              : { label: isDe ? 'Plätze' : 'Seats', value: unlimitedParticipants ? (isDe ? 'Unbegrenzt' : 'Unlimited') : String(maxParticipants || 0), status: unlimitedParticipants ? 'default' : 'ok' },
            { label: isDe ? 'Warteliste' : 'Waitlist', value: waitlistEnabled ? (isDe ? 'aktiv' : 'on') : (isDe ? 'aus' : 'off'), status: waitlistEnabled ? 'default' : 'ok' },
            { label: isDe ? 'Anmeldefrist' : 'Registration deadline', value: registrationDeadline ? fmtDt(registrationDeadline) : '—', status: registrationDeadline ? 'ok' : 'empty' },
            { label: isDe ? 'Abmeldefrist (kommuniziert)' : 'Cancellation deadline (communicated)', value: lastDeregisterDate ? fmtDt(lastDeregisterDate) : '—', status: lastDeregisterDate ? 'ok' : 'empty' },
            { label: isDe ? 'Sichtbarkeit' : 'Visibility', value: (locList.length === 0 && audList.length === 0) ? (isDe ? 'alle Mitarbeiter von Deloitte Deutschland' : 'all Deloitte Germany employees') : [locList.length ? `${isDe ? 'Standorte' : 'Locations'}: ${locList.join(', ')}` : '', audList.length ? `${audList.length} ${isDe ? 'Verteiler/Personen' : 'lists/people'}` : ''].filter(Boolean).join(filterMode === 'AND' ? ' UND ' : ' ODER '), status: (locList.length === 0 && audList.length === 0) ? 'default' : 'ok' },
            { label: isDe ? 'Ausgeschlossene Personen' : 'Excluded people', value: excludedUsers.length ? `${excludedUsers.length}` : '—', status: excludedUsers.length ? 'ok' : 'default' },
          ],
        });
        sections.push({
          title: isDe ? 'Schritt 5 — Felder' : 'Step 5 — Fields',
          rows: [
            { label: isDe ? 'Eigene Abfrage-Felder' : 'Custom fields', value: customFields.length ? `${customFields.length}` : (isDe ? 'keine' : 'none'), status: customFields.length ? 'ok' : 'default' },
            { label: isDe ? 'Anrede abfragen' : 'Ask salutation', value: askSalutation ? (isDe ? 'an' : 'on') : (isDe ? 'aus' : 'off'), status: askSalutation ? 'ok' : 'default' },
            { label: isDe ? 'Zweisprachige Felder (DE+EN)' : 'Bilingual fields (DE+EN)', value: bilingualFields ? (isDe ? 'an' : 'on') : (isDe ? 'aus' : 'off'), status: bilingualFields ? 'ok' : 'default' },
            { label: isDe ? 'Formular-Sprache' : 'Form language', value: registrationLanguage === 'de' ? (isDe ? 'Immer Deutsch' : 'Always German') : registrationLanguage === 'en' ? (isDe ? 'Immer Englisch' : 'Always English') : (isDe ? 'Automatisch (App-Sprache)' : 'Automatic (app language)'), status: registrationLanguage ? 'ok' : 'default' },
          ],
        });
        sections.push({
          title: isDe ? 'Schritt 6 — Kommunikation' : 'Step 6 — Communication',
          rows: [
            { label: isDe ? 'Mail-Sprache' : 'Email language', value: (topComm.emailLanguage || 'EN').toUpperCase() === 'DE' ? 'Deutsch' : 'English', status: 'ok' },
            { label: isDe ? 'Bestätigungs-Mails' : 'Confirmation emails', value: topComm.disableEmails ? (isDe ? 'deaktiviert' : 'disabled') : (isDe ? 'aktiv' : 'on'), status: topComm.disableEmails ? 'ok' : 'default' },
            ...(!topComm.disableEmails && (topComm.disableRegistrationEmail || topComm.disableCancellationEmail) ? [{ label: isDe ? 'Einzeln deaktiviert' : 'Individually disabled', value: [topComm.disableRegistrationEmail ? (isDe ? 'Anmelde-Bestätigung' : 'registration confirmation') : '', topComm.disableCancellationEmail ? (isDe ? 'Abmelde-Bestätigung' : 'cancellation confirmation') : ''].filter(Boolean).join(', '), status: 'ok' as CheckStatus }] : []),
            { label: isDe ? 'Outlook-Termin' : 'Outlook invite', value: topComm.disableOutlook ? (isDe ? 'deaktiviert' : 'disabled') : (isDe ? 'aktiv' : 'on'), status: topComm.disableOutlook ? 'ok' : 'default' },
            { label: isDe ? 'Auto-Abmeldung bei Outlook-Absage' : 'Auto-cancel on Outlook decline', value: topComm.autoDeregisterOnDecline ? (isDe ? 'an' : 'on') : (isDe ? 'aus' : 'off'), status: topComm.autoDeregisterOnDecline ? 'ok' : 'default' },
            { label: isDe ? 'Person nicht mehr bei Deloitte' : 'Person no longer at Deloitte', value: topComm.inactiveHandling === 'autoderegister' ? (isDe ? 'automatisch abmelden' : 'auto-deregister') : (isDe ? 'Organizer informieren' : 'notify organizer'), status: topComm.inactiveHandling === 'autoderegister' ? 'ok' : 'default' },
          ],
        });
        sections.push({
          title: isDe ? 'Schritt 7 — Team-Anmeldung' : 'Step 7 — Team registration',
          rows: [{ label: isDe ? 'Team-Anmeldung' : 'Team registration', value: teamRegistrationEnabled ? (isDe ? `aktiv — Teams à ${teamSize}` : `on — teams of ${teamSize}`) : (isDe ? 'aus' : 'off'), status: teamRegistrationEnabled ? 'ok' : 'default' }],
        });
        sections.push({
          title: isDe ? 'Schritt 8 — Dokumente' : 'Step 8 — Documents',
          rows: [{ label: isDe ? 'Dokumente' : 'Documents', value: documents.length ? `${documents.length}` : '—', status: documents.length ? 'ok' : 'empty' }],
        });
        sections.push({
          title: isDe ? 'Schritt 9 — Fun-Zone' : 'Step 9 — Fun zone',
          rows: [{ label: 'Quiz', value: quiz.length ? `${quiz.length} ${isDe ? 'Fragen' : 'questions'}` : '—', status: quiz.length ? 'ok' : 'empty' }],
        });
        const missingCount = sections.reduce((acc, s) => acc + s.rows.filter(r => r.status === 'missing').length, 0);
        const emptyCount = sections.reduce((acc, s) => acc + s.rows.filter(r => r.status === 'empty').length, 0);
        // v31.2: Das Ergebnis steht zuerst und nennt die fehlenden Angaben
        // beim Namen — vorher musste man neun Abschnitte nach roten Chips
        // absuchen, um zu wissen, WAS fehlt.
        const missingLabels: string[] = [];
        sections.forEach(s => s.rows.forEach(r => { if (r.status === 'missing') missingLabels.push(r.label); }));
        const chip = (st: CheckStatus): React.ReactElement | null => {
          if (st === 'default') return <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? 'Standard' : 'Default'}</span>;
          if (st === 'empty') return <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? 'leer (optional)' : 'empty (optional)'}</span>;
          if (st === 'missing') return <span className="dex-ui-pill dex-ui-pill--red">{isDe ? 'fehlt' : 'missing'}</span>;
          return <span className="dex-ui-pill dex-ui-pill--green" aria-label={isDe ? 'gesetzt' : 'set'}><Check size={12} /></span>;
        };
        const verdictClass = missingCount > 0 ? 'dex-ui-callout--danger' : emptyCount > 0 ? 'dex-ui-callout--warn' : 'dex-ui-callout--success';
        return (
          <Modal
            open={true}
            onClose={() => setShowConfigCheck(false)}
            maxWidth={760}
            ariaLabel={isDe ? 'Event prüfen' : 'Review event'}
            icon={<Check size={20} />}
            title={isDe ? 'Event prüfen' : 'Review event'}
            subtitle={isDe
              ? 'Alle Einstellungen im Überblick — was gesetzt ist, wo Standards greifen und was noch fehlt. Geändert wird hier nichts.'
              : 'All settings at a glance — what is set, where defaults apply and what is still missing. Nothing is changed here.'}
            footer={(
              <button type="button" className="btn btn-primary" onClick={() => setShowConfigCheck(false)}>{isDe ? 'Schließen' : 'Close'}</button>
            )}
          >
            <div className={cx('dex-ui-callout', verdictClass)} role="status">
              <span className="dex-ui-callout-icon">{missingCount > 0 ? <AlertCircle size={16} /> : <Check size={16} />}</span>
              <span>
                {missingCount > 0
                  ? (isDe
                    ? <><strong>{missingCount === 1 ? '1 Pflichtangabe fehlt' : `${missingCount} Pflichtangaben fehlen`}:</strong> {missingLabels.join(', ')}.{emptyCount > 0 ? <> Dazu {emptyCount === 1 ? 'ist 1 optionaler Punkt' : `sind ${emptyCount} optionale Punkte`} noch leer.</> : null}</>
                    : <><strong>{missingCount === 1 ? '1 required item missing' : `${missingCount} required items missing`}:</strong> {missingLabels.join(', ')}.{emptyCount > 0 ? <> Plus {emptyCount} optional {emptyCount === 1 ? 'item is' : 'items are'} still empty.</> : null}</>)
                  : emptyCount > 0
                    ? (isDe ? <>Alle Pflichtangaben sind gesetzt — <strong>{emptyCount === 1 ? '1 optionaler Punkt ist' : `${emptyCount} optionale Punkte sind`}</strong> noch leer.</> : <>All required items are set — <strong>{emptyCount} optional {emptyCount === 1 ? 'item is' : 'items are'}</strong> still empty.</>)
                    : (isDe ? 'Alles gesetzt — keine offenen Punkte.' : 'Everything set — nothing open.')}
              </span>
            </div>
            <div>
              {sections.map((sec, si) => (
                <div key={si} className="dex-ui-section">
                  <div className="dex-ui-section-title">{sec.title}</div>
                  <div className="dex-ui-card" style={{ padding: '2px 8px' }}>
                    {sec.rows.map((r, ri) => (
                      <div key={ri} className="dex-ui-row dex-ui-row--bordered" style={{ flexDirection: isMobile ? 'column' : 'row', alignItems: isMobile ? 'stretch' : 'center', gap: isMobile ? 4 : 10, padding: '8px 6px', fontSize: '0.82rem' }}>
                        <span style={{ flex: isMobile ? '0 0 auto' : '0 0 220px', color: 'var(--dex-gray-500)' }}>{r.label}</span>
                        <span style={{ flex: 1, color: 'var(--dex-gray-800)', minWidth: 0, overflowWrap: 'anywhere' }}>{r.value}</span>
                        {chip(r.status)}
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </Modal>
        );
      })()}

      {showDemoVariantModal && (() => {
        const cards: Array<{ key: keyof typeof DEMO_VARIANTS; titleDe: string; titleEn: string; descDe: string; descEn: string }> = [
          {
            key: 'standard',
            titleDe: 'Standard',
            titleEn: 'Standard',
            descDe: 'Ein Event, eine Gruppe — typisches Meeting oder Lunch.',
            descEn: 'One event, one group — a typical meeting or lunch.',
          },
          {
            key: 'groups',
            titleDe: 'Mit Gruppen',
            titleEn: 'With groups',
            descDe: 'Zwei Teilnehmer-Gruppen mit geteilter Kapazität, z.B. Vormittag / Nachmittag.',
            descEn: 'Two participant groups with split capacity, e.g. morning / afternoon.',
          },
          {
            key: 'subevent',
            titleDe: 'Mit Sub-Event',
            titleEn: 'With sub-event',
            descDe: 'Haupt-Event + 1 Sub-Event, z.B. Konferenz + Dinner.',
            descEn: 'Main event + 1 sub-event, e.g. conference + dinner.',
          },
          {
            key: 'subeventTeam',
            titleDe: 'Mit Sub-Event + Team',
            titleEn: 'With sub-event + team',
            // v31.2: „Wie links" stimmt im Raster nicht mehr, sobald die
            // Kacheln umbrechen — deshalb die Vorlage beim Namen nennen.
            descDe: 'Wie „Mit Sub-Event“, zusätzlich Team-Anmeldung (Teams à 4 Personen).',
            descEn: 'Like “With sub-event”, plus team registration (teams of 4 people).',
          },
        ];
        return (
          <Modal
            open={true}
            onClose={() => setShowDemoVariantModal(false)}
            maxWidth={720}
            ariaLabel={isDe ? 'Demo-Daten laden' : 'Load demo data'}
            icon={<Star size={20} />}
            title={isDe ? 'Demo-Daten laden' : 'Load demo data'}
            subtitle={isDe
              ? 'Wähle eine Vorlage — sie überschreibt deine aktuellen Eingaben. Solange du noch nichts gespeichert hast, geht dabei nichts verloren.'
              : 'Pick a template — it overrides your current input. As long as you haven’t saved yet, nothing is lost.'}
            footer={(
              <button type="button" className="btn btn-secondary" onClick={() => setShowDemoVariantModal(false)}>
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
            )}
          >
            <div className="dex-ui-grid-2">
              {cards.map(card => (
                <button
                  key={card.key}
                  type="button"
                  className="dex-ui-choice"
                  onClick={() => {
                    DEMO_VARIANTS[card.key]();
                    setShowDemoVariantModal(false);
                  }}
                >
                  <span className="dex-ui-choice-body">
                    <span className="dex-ui-choice-title" style={{ display: 'block' }}>{isDe ? card.titleDe : card.titleEn}</span>
                    <span className="dex-ui-choice-desc" style={{ display: 'block' }}>{isDe ? card.descDe : card.descEn}</span>
                  </span>
                </button>
              ))}
            </div>
          </Modal>
        );
      })()}

      {/* v17.21: A4-Zusammenfassungs-Modal nach erfolgreichem Save — fragt
          den Organizer, ob er das gesamte Event als PDF oder Word herunter-
          laden möchte (z.B. zur Durchsicht durch einen Partner). Beim
          Klick auf eine Option läuft der Export sofort, danach feuert der
          eigentliche „Wizard verlassen"-Dispatch (`dex-event-submit-success`).
          „Nein, danke" springt direkt zum Dispatch. */}
      {/* v28.74: „Einstellungen auf andere übertragen" — Auswahl WAS und WOHIN. */}
      {subTransfer && (() => {
        const srcName = shortSubEventTitle(subEvents[subTransfer.fromIdx]?.title || '', title) || (childTermSingular || 'Sub-Event');
        const toggleGroup = (key: string): void => setSubTransfer(prev => prev && ({
          ...prev,
          groups: prev.groups.indexOf(key) >= 0 ? prev.groups.filter(k => k !== key) : [...prev.groups, key],
        }));
        const toggleTarget = (i: number): void => setSubTransfer(prev => prev && ({
          ...prev,
          targets: prev.targets.indexOf(i) >= 0 ? prev.targets.filter(x => x !== i) : [...prev.targets, i],
        }));
        const canApply = subTransfer.groups.length > 0 && subTransfer.targets.length > 0;
        const allTargetsOn = subTransfer.targets.length === subEvents.length - 1;
        return (
          <Modal
            open={true}
            onClose={() => setSubTransfer(null)}
            maxWidth={680}
            // v31.2: Wie vorher schließt ein Klick neben die Karte NICHT —
            // sonst ist die Auswahl weg. Escape und X wirken wie „Abbrechen".
            backdropClose={false}
            ariaLabel={isDe ? 'Einstellungen übertragen' : 'Transfer settings'}
            icon={<Copy size={20} />}
            title={isDe ? 'Einstellungen übertragen' : 'Transfer settings'}
            subtitle={isDe
              ? <>Von <strong>&bdquo;{srcName}&ldquo;</strong> auf andere {childTermPlural || 'Sub-Events'} — die Werte dort werden <strong>überschrieben</strong>. Gespeichert wird erst, wenn du den Assistenten speicherst.</>
              : <>From <strong>&bdquo;{srcName}&ldquo;</strong> to other sub-events — their values are <strong>overwritten</strong>. Nothing is stored until you save the wizard.</>}
            footer={<>
              <button type="button" className="btn btn-secondary" onClick={() => setSubTransfer(null)}>
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button type="button" className="btn btn-primary" disabled={!canApply} onClick={applySubTransfer}>
                {isDe
                  ? `Auf ${subTransfer.targets.length} übertragen`
                  : `Transfer to ${subTransfer.targets.length}`}
              </button>
            </>}
          >
            <div>
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? '1 · Was soll übertragen werden?' : '1 · What should be transferred?'}</div>
                <div className="dex-ui-stack" style={{ gap: 6 }}>
                  {SUB_TRANSFER_GROUPS.map(g => {
                    const on = subTransfer.groups.indexOf(g.key) >= 0;
                    const n = subGroupDiffCount(subTransfer.fromIdx, g.fields);
                    return (
                      <label key={g.key} className={cx('dex-ui-toggle-row', on && 'is-active')} style={{ padding: '9px 12px' }}>
                        <input type="checkbox" checked={on} onChange={() => toggleGroup(g.key)} />
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">
                            {isDe ? g.de : g.en}
                            {n > 0 && (
                              <span className="dex-ui-pill dex-ui-pill--orange" title={isDe ? 'So viele Ziel-Termine haben hier andere Werte' : 'That many target sessions have different values here'}>
                                {isDe ? `${n}× abweichend` : `${n}× differing`}
                              </span>
                            )}
                          </span>
                        </span>
                      </label>
                    );
                  })}
                </div>
              </div>

              <div className="dex-ui-section">
                <div className="dex-ui-inline" style={{ justifyContent: 'space-between', marginBottom: 10 }}>
                  <div className="dex-ui-section-title" style={{ flex: 1, margin: 0 }}>{isDe ? '2 · Auf welche übertragen?' : '2 · Transfer to which ones?'}</div>
                  <button
                    type="button"
                    className="dex-ui-textbtn"
                    onClick={() => setSubTransfer(prev => prev && ({
                      ...prev,
                      targets: prev.targets.length === subEvents.length - 1
                        ? []
                        : subEvents.map((_, i) => i).filter(i => i !== prev.fromIdx),
                    }))}
                  >
                    {allTargetsOn
                      ? (isDe ? 'Keine auswählen' : 'Select none')
                      : (isDe ? 'Alle auswählen' : 'Select all')}
                  </button>
                </div>
                <div className="dex-ui-inline" style={{ gap: 6 }}>
                  {subEvents.map((s, i) => {
                    if (i === subTransfer.fromIdx) return null;
                    const on = subTransfer.targets.indexOf(i) >= 0;
                    const nm = shortSubEventTitle(s.title, title) || (isDe ? 'Ohne Titel' : 'Untitled');
                    return (
                      <button key={s.id || i} type="button" className={cx('dex-ui-chip', on && 'is-active')} aria-pressed={on} onClick={() => toggleTarget(i)}>
                        {on && <Check size={12} />}
                        {nm}
                      </button>
                    );
                  })}
                </div>
              </div>
            </div>
          </Modal>
        );
      })()}

      {showSummaryModal && pendingSuccessDispatch && (() => {
        const closeAndDispatch = (): void => {
          // v17.22: Ref VOR dem Dispatch leeren, damit der Unmount-Cleanup
          // nicht ein zweites Mal feuert (Doppel-Navigation/-Banner).
          const payload = pendingSuccessDispatchRef.current || pendingSuccessDispatch;
          pendingSuccessDispatchRef.current = null;
          setShowSummaryModal(false);
          setPendingSuccessDispatch(null);
          try {
            window.dispatchEvent(new CustomEvent('dex-event-submit-success', {
              detail: payload,
            }));
          } catch { /* */ }
        };
        const buildData = (): SummaryData => {
          // Bild als DataURL (falls noch nicht Base64): unten reicht der
          // bestehende imagePreview, der bei neu hochgeladenen Bildern
          // bereits eine Data-URL ist und bei bestehenden Events die
          // SharePoint-URL. Letztere wird im PDF/Doc-Export im Print-View
          // i.d.R. nicht geladen (CORS) — wir bauen einen Fallback-Text.
          const subEventsForSummary = subEvents.map(se => ({
            title: se.title || '',
            startDate: se.startDate,
            endDate: se.endDate,
            location: se.location,
            description: se.description,
            maxParticipants: typeof se.maxParticipants === 'number' ? se.maxParticipants : undefined,
            waitlistEnabled: !!se.waitlistEnabled,
          }));
          const customFieldsForSummary = customFields
            .filter(f => f.label && f.label.trim().length > 0)
            .map(f => ({
              id: f.id,
              label: f.label,
              type: f.type,
              required: !!f.required,
              helpText: f.helpText,
              helpTextStyle: f.helpTextStyle,
              confirmLabel: f.confirmLabel,
              options: f.type === 'select' ? f.options : undefined,
              multi: !!f.multi,
              onlyForGroup: f.onlyForGroup,
              labelEn: f.labelEn,
              helpTextEn: f.helpTextEn,
              confirmLabelEn: f.confirmLabelEn,
              optionsEn: f.optionsEn,
              showIf: f.showIf,
            }));
          // Transferzeiten + Agenda werden in den Summary-Helper als
          // vereinfachte Spalten gemappt — das Detail-Schema bleibt im
          // Wizard, der Export nimmt die für Reviewer relevanten Spalten.
          const transfersForSummary = transferTimes.map(t => ({
            time: [t.date, t.departureTime].filter(Boolean).join(' '),
            description: [t.location, t.description, t.meetingPoint].filter(Boolean).join(' — '),
          }));
          const agendaForSummary = agenda.map(a => ({
            time: [a.date, a.time, a.endTime ? ` – ${a.endTime}` : ''].filter(Boolean).join(' '),
            topic: a.title,
            speaker: a.description,
          }));
          const quizForSummary = quiz.map(q => ({
            question: q.question,
            options: q.options,
            correctIndex: (q.correctIndices && q.correctIndices.length > 0) ? q.correctIndices[0] : undefined,
          }));
          const documentsForSummary = documents.map(doc => ({
            name: doc.name,
            size: doc.size,
          }));
          return {
            title,
            description,
            imageDataUrl: imagePreview || eventImageUrl || undefined,
            startDate,
            endDate,
            organizers: organizer.split(';').map(s => s.trim()).filter(Boolean),
            organizerEmails,
            contactName,
            contactEmail,
            contactInfo,
            testTeam: testTeamNames,
            qrScanners: qrScannerNames,
            isFictive,
            activeFrom,
            location,
            address: { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity },
            agenda: agendaForSummary,
            transfers: transfersForSummary,
            // v29.21 (Audit): beide Strings sind im Wizard KOMMA-separiert —
            // split(';') lieferte immer ein 1-elementiges Array, der Export
            // meldete bei fuenf Verteilern „1 Eintrag".
            locationFilter: locationFilter ? locationFilter.split(',').map(s => s.trim()).filter(Boolean) : [],
            audience: audience ? audience.split(',').map(s => s.trim()).filter(Boolean) : [],
            filterMode,
            excludedUsers,
            registrationDeadline,
            lastDeregisterDate,
            maxParticipants: Number(maxParticipants) || 0,
            unlimitedParticipants,
            waitlistEnabled,
            durchstarterCapacity: Number(durchstarterCapacity) || 0,
            funstarterCapacity: Number(funstarterCapacity) || 0,
            splitLabelA, splitLabelB,
            splitSharedWaitlist,
            teamRegistrationEnabled,
            teamSize,
            askTeamName,
            teamPartialAllowed,
            teamOpenSlotsVisible,
            teamJoinRequiresApproval,
            askSalutation,
            bilingualFields,
            customFields: customFieldsForSummary,
            allowAttendeeUpload,
            attendeeUploadHint,
            attendeeUploadLabel,
            emailLanguage,
            disableEmails,
            disableOutlook,
            outlookHeading,
            outlookSubheading,
            outlookBody,
            notifyOrgRegisterMode,
            notifyOrgRegisterFromDate,
            notifyOrgCancelMode,
            documents: documentsForSummary,
            funZone: quizForSummary,
            quizClusterSize,
            subEvents: subEventsForSummary,
            childTermSingular,
            childTermPlural,
            subEventsOnlyMode,
            requireSubEventSelection,
            generatedAt: new Date().toISOString(),
            locale: isDe ? 'de' : 'en',
          };
        };
        const onPdf = (): void => {
          try { exportSummaryAsPdf(buildData()); } catch (err) {
            console.warn('[DEX] exportSummaryAsPdf failed:', err);
          }
          closeAndDispatch();
        };
        const onDoc = (): void => {
          try { exportSummaryAsDoc(buildData()); } catch (err) {
            console.warn('[DEX] exportSummaryAsDoc failed:', err);
          }
          closeAndDispatch();
        };
        return (
          <Modal
            open={true}
            onClose={closeAndDispatch}
            maxWidth={560}
            ariaLabel={isDe ? 'Event-Zusammenfassung herunterladen' : 'Download event summary'}
            icon={<FileText size={20} />}
            title={isDe ? 'Gespeichert — Zusammenfassung herunterladen?' : 'Saved — download a summary?'}
            subtitle={isDe
              ? 'Eine A4-Seite mit allen Angaben zum Event — zum Beispiel für einen Partner zur Durchsicht.'
              : 'A one-page A4 overview of the whole event — for example for a partner to review.'}
            // v31.2: „Nein, danke" links mit Abstand, die beiden Export-Wege
            // rechts; PDF ist der Hauptweg.
            footer={<>
              <div className="dex-ui-modal-foot-left">
                <button type="button" className="btn btn-secondary" onClick={closeAndDispatch}>
                  {isDe ? 'Nein, danke' : 'No, thanks'}
                </button>
              </div>
              <button type="button" className="btn btn-outline" onClick={onDoc}>
                <Download size={14} /> {isDe ? 'Als Word (.doc)' : 'As Word (.doc)'}
              </button>
              <button type="button" className="btn btn-primary" onClick={onPdf}>
                <Download size={14} /> {isDe ? 'Als PDF' : 'As PDF'}
              </button>
            </>}
          >
            <div className="dex-ui-stack">
              <div className="dex-ui-callout dex-ui-callout--success">
                <span className="dex-ui-callout-icon"><Check size={16} /></span>
                <span>
                  {isDe
                    ? <>Das Event wurde gespeichert. Die Zusammenfassung enthält <strong>alle Sektionen</strong>: Foto, Beschreibung, Sichtbarkeit, Felder, Kommunikation, Dokumente, Sub-Events …</>
                    : <>The event has been saved. The summary contains <strong>every section</strong>: photo, description, visibility, fields, communication, documents, sub-events …</>}
                </span>
              </div>
              <div className="dex-ui-callout dex-ui-callout--info">
                <span className="dex-ui-callout-icon"><Info size={16} /></span>
                <span>
                  {isDe
                    ? <><strong>PDF:</strong> Der Browser-Druckdialog öffnet sich — wähle dort <strong>&bdquo;Als PDF speichern&ldquo;</strong> als Ziel. <strong>Word:</strong> lädt direkt eine .doc-Datei herunter.</>
                    : <><strong>PDF:</strong> the browser print dialog opens — pick <strong>&ldquo;Save as PDF&rdquo;</strong> as the destination. <strong>Word:</strong> downloads a .doc file directly.</>}
                </span>
              </div>
            </div>
          </Modal>
        );
      })()}

      {/* v19.x: Das „Personen ausschließen"-Modal UND das „Sichtbarkeit prüfen"-Modal
          sind nach <AudiencePicker> gewandert (self-contained pro Instanz für
          Hauptevent + jedes Sub-Event). */}

      {/* v20.2: Self-Check-in-Erklär-Modal (v18.33) entfernt — die Erklärung
          lebt jetzt im Kachel-Modal des Admin Centers und im Handbuch. */}

      {/* v9.28/v13.4: Modal — neuer Quiz-Bereich anlegen, jetzt über <Modal>-Wrapper. */}
      <Modal
        open={newSectionModalOpen}
        onClose={() => setNewSectionModalOpen(false)}
        maxWidth={460}
        // v31.2: Eingabe-Dialog — ein Klick neben die Karte soll den
        // getippten Namen nicht wegwerfen (Escape, X und Abbrechen schließen).
        backdropClose={false}
        ariaLabel="Neuen Quiz-Bereich anlegen"
        icon={<Plus size={20} />}
        title={isDe ? 'Neuen Bereich anlegen' : 'Create a new section'}
        subtitle={isDe
          ? 'Ein Bereich bündelt Quiz-Fragen auf einer gemeinsamen Seite.'
          : 'A section groups quiz questions on one shared page.'}
        footer={<>
          <button type="button" className="btn btn-secondary" onClick={() => setNewSectionModalOpen(false)}>
            {isDe ? 'Abbrechen' : 'Cancel'}
          </button>
          <button type="button" className="btn btn-primary" onClick={submitNewSection}>
            <Plus size={14} /> {isDe ? 'Bereich anlegen' : 'Create section'}
          </button>
        </>}
      >
        {newSectionModalOpen && (
          <div className="dex-ui-field">
            <label className="dex-ui-label" htmlFor="dex-new-section-name">
              {isDe ? 'Wie soll der Bereich heißen?' : 'What should the section be called?'}
            </label>
            <input
              id="dex-new-section-name"
              type="text"
              className="dex-ui-input"
              autoFocus
              value={newSectionName}
              placeholder={isDe ? 'z.B. Orte, Geschichte, Foto-Quiz' : 'e.g. Places, History, Photo quiz'}
              aria-invalid={!!newSectionError}
              onChange={e => { setNewSectionName(e.target.value); setNewSectionError(''); }}
              onKeyDown={e => {
                if (e.key === 'Enter') {
                  submitNewSection();
                } else if (e.key === 'Escape') {
                  setNewSectionModalOpen(false);
                }
              }}
            />
            {newSectionError ? (
              <div className="dex-ui-help" role="alert" style={{ color: 'var(--dex-red, #c00)', fontWeight: 600 }}>{newSectionError}</div>
            ) : (
              <div className="dex-ui-help">
                {isDe ? 'Kurz und sprechend — der Name steht als Überschrift über den Fragen des Bereichs.' : 'Short and descriptive — the name appears as the heading above the section’s questions.'}
              </div>
            )}
          </div>
        )}
      </Modal>

      {/* Modal: Vorgeschlagene Felder auswählen (Multi-Select) — v13.4 über <Modal>. */}
      {(() => {
        const suggestedCount = Object.values(suggestedSelection).filter(Boolean).length;
        return (
          <Modal
            open={showSuggestedModal}
            onClose={() => setShowSuggestedModal(false)}
            maxWidth={560}
            ariaLabel="Vorgeschlagene Felder auswählen"
            icon={<Plus size={20} />}
            title={isDe ? 'Vorgeschlagene Felder' : 'Suggested fields'}
            subtitle={isDe
              ? 'Welche Fragen soll dein Anmeldeformular zusätzlich stellen? Anpassen kannst du jedes Feld danach weiter.'
              : 'Which extra questions should your registration form ask? You can still tweak every field afterwards.'}
            // v31.2: Sammel-Aktionen links als Textknöpfe, Entscheidung rechts —
            // der Primär-Knopf zählt mit, damit klar ist, was gleich passiert.
            footer={<>
              <div className="dex-ui-modal-foot-left">
                <button
                  type="button"
                  className="dex-ui-textbtn"
                  onClick={() => {
                    const all: Record<string, boolean> = {};
                    for (const s of SUGGESTED_FIELDS_CATALOG) all[s.key] = true;
                    setSuggestedSelection(all);
                  }}
                >
                  {isDe ? 'Alle auswählen' : 'Select all'}
                </button>
                <button
                  type="button"
                  className="dex-ui-textbtn dex-ui-textbtn--muted"
                  onClick={() => setSuggestedSelection({})}
                  disabled={suggestedCount === 0}
                >
                  {isDe ? 'Auswahl aufheben' : 'Clear selection'}
                </button>
              </div>
              <button type="button" className="btn btn-secondary" onClick={() => setShowSuggestedModal(false)}>
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={addSelectedSuggestedFields}
                disabled={suggestedCount === 0}
              >
                <Plus size={14} /> {suggestedCount > 0
                  ? (isDe ? `${suggestedCount} ${suggestedCount === 1 ? 'Feld' : 'Felder'} hinzufügen` : `Add ${suggestedCount} ${suggestedCount === 1 ? 'field' : 'fields'}`)
                  : (isDe ? 'Hinzufügen' : 'Add')}
              </button>
            </>}
          >
            {/* v10.21: Catalog gruppiert nach Kategorie. Allgemeine Felder
                immer ausgeklappt, B2Run-Felder default eingeklappt mit
                Toggle. Jeder Eintrag bekommt ein Badge mit der Kategorie. */}
            {showSuggestedModal && (() => {
              const generalEntries = SUGGESTED_FIELDS_CATALOG.filter(s => s.category === 'general');
              const b2runEntries = SUGGESTED_FIELDS_CATALOG.filter(s => s.category === 'b2run');
              const renderEntry = (s: SuggestedEntry): React.ReactElement => (
                <label key={s.key} className={cx('dex-ui-toggle-row', !!suggestedSelection[s.key] && 'is-active')}>
                  <input
                    type="checkbox"
                    checked={!!suggestedSelection[s.key]}
                    onChange={e => setSuggestedSelection({ ...suggestedSelection, [s.key]: e.target.checked })}
                  />
                  {/* v10.23: passendes Fluent-UI-Icon links neben dem Label,
                      damit die Auswahl auf einen Blick visuell wiedererkennbar
                      ist. Farbe analog zur Kategorie (grün=Allgemein,
                      orange=B2Run). */}
                  <Icon
                    iconName={s.icon}
                    style={{
                      fontSize: 20, flexShrink: 0, marginTop: 1,
                      color: s.category === 'b2run' ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-green-dark, #4a7c1f)',
                    }}
                  />
                  <span className="dex-ui-toggle-row-body">
                    <span className="dex-ui-toggle-row-title">
                      {s.label}
                      <span className={cx('dex-ui-pill', s.category === 'b2run' ? 'dex-ui-pill--orange' : 'dex-ui-pill--green')}>
                        {s.category === 'b2run' ? 'B2Run' : (isDe ? 'Allgemein' : 'General')}
                      </span>
                      {/* v10.23: i-Tooltip mit ausführlichem Hinweis was das
                          Feld in der App tut — verhindert Klick-und-Probier-
                          Modus, weil der Organizer schon vor Auswahl sieht
                          welche Frage-Form (Dropdown / Freitext / Pflicht-
                          Checkbox) und welcher Effekt (Anzeige im Admin-Center,
                          Excel-Export, etc.) entsteht. Klick auf das Label
                          (das `<label>`-Wrapping) würde die Checkbox togglen
                          — das `onClick`-stopPropagation des InfoTooltip
                          verhindert das. */}
                      <span onClick={e => e.preventDefault()} style={{ display: 'inline-flex' }}>
                        <InfoTooltip text={s.tooltip || s.description} />
                      </span>
                    </span>
                    <span className="dex-ui-toggle-row-desc" style={{ display: 'block' }}>{s.description}</span>
                  </span>
                </label>
              );
              return (
                <div className="dex-ui-stack">
                  {generalEntries.map(renderEntry)}
                  <div>
                    <button
                      type="button"
                      className={cx('dex-ui-disclosure', showB2runSuggested && 'is-open')}
                      aria-expanded={showB2runSuggested}
                      onClick={() => setShowB2runSuggested(v => !v)}
                    >
                      <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                      {isDe ? 'B2Run-spezifische Felder' : 'B2Run-specific fields'}
                      <span className="dex-ui-pill dex-ui-pill--orange">B2Run</span>
                      <span className="dex-ui-disclosure-count">{b2runEntries.length}</span>
                    </button>
                    {showB2runSuggested && (
                      <div className="dex-ui-disclosure-body dex-ui-stack">
                        <p className="dex-ui-muted" style={{ margin: 0 }}>
                          {isDe
                            ? 'Nur für B2Run-Lauf-Events (Startblock, Altersklasse, Datenschutz-Checkbox mit b2run.de-Links). Bei normalen Events brauchst du sie nicht.'
                            : 'Only for B2Run running events (start block, age group, B2Run-specific privacy checkbox). Skip them for standard events.'}
                        </p>
                        {b2runEntries.map(renderEntry)}
                      </div>
                    )}
                  </div>
                </div>
              );
            })()}
          </Modal>
        );
      })()}

      {/* v22.62: „Sichtbarkeit auf Sub-Events übernehmen?" — erscheint beim
          ersten „Weiter"/Speichern, sobald die Klammer eine Sichtbarkeit hat
          und Sub-Events existieren. */}
      <Modal
        open={visCopyModalOpen}
        onClose={() => closeVisCopy(false)}
        maxWidth={560}
        dismissable={false}
        hideClose
        ariaLabel={isDe ? 'Sichtbarkeit übernehmen' : 'Apply visibility'}
        icon={<Users size={20} />}
        title={isDe ? 'Sichtbarkeit auf alle Sub-Events übernehmen?' : 'Apply visibility to all sub-events?'}
        subtitle={isDe
          ? <>Du hast für {subEventsOnlyMode ? 'die Klammer' : 'das Hauptevent'} eine Sichtbarkeit gesetzt. Sollen <strong>alle {subEvents.length} Sub-Events</strong> dieselbe übernehmen — Standortfilter, Mailverteiler und Verknüpfung?</>
          : <>You set a visibility for {subEventsOnlyMode ? 'the bracket' : 'the main event'}. Should <strong>all {subEvents.length} sub-events</strong> adopt the same one — location filter, mailing lists and combination?</>}
        footer={<>
          <button type="button" className="btn btn-secondary" onClick={() => closeVisCopy(false)}>
            {isDe ? 'Nein, eigene behalten' : 'No, keep their own'}
          </button>
          <button type="button" className="btn btn-primary" onClick={() => closeVisCopy(true)}>
            {isDe ? `Ja, auf alle ${subEvents.length} übernehmen` : `Yes, apply to all ${subEvents.length}`}
          </button>
        </>}
      >
        {visCopyModalOpen && (
          <div className="dex-ui-stack">
            <div className="dex-ui-callout dex-ui-callout--info">
              <span className="dex-ui-callout-icon"><Info size={16} /></span>
              <span>
                {isDe
                  ? 'Meist sinnvoll: Wer das Event sehen soll, erreicht so auch die Sub-Events — der Zugang läuft ohnehin über die Sichtbarkeit des Gesamt-Events.'
                  : 'Usually a good idea: everyone who should see the event can then also reach the sub-events — access runs through the overall event’s visibility anyway.'}
              </span>
            </div>
            <div className="dex-ui-callout dex-ui-callout--warn">
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <span>
                {isDe
                  ? <>Bereits gesetzte, abweichende Sub-Event-Sichtbarkeiten werden dabei <strong>überschrieben</strong>.</>
                  : <>Any existing, differing sub-event visibilities will be <strong>overwritten</strong>.</>}
              </span>
            </div>
          </div>
        )}
      </Modal>

      {/* v11.57 / v11.63 / v13.4: Outlook-Update-Confirm-Modal über <Modal>-Wrapper.
          dismissable=false, da Schließen nur über Cancel-Button erlaubt. */}
      {(() => {
        // v31.2: Der Primär-Knopf nennt die Folge — „nur speichern" oder
        // „speichern und N Termine aktualisieren" — statt eines nackten
        // „Speichern", bei dem unklar bleibt, ob jetzt Mails rausgehen.
        const checkedCount = outlookConfirmItems.filter(it => !!outlookConfirmChecks[it.eventId]).length;
        const fieldLabelMap: Record<'title'|'startDate'|'endDate'|'outlookBody'|'location'|'subject'|'layout'|'organizer'|'logo', { de: string; en: string }> = {
          title: { de: 'Titel', en: 'Title' },
          startDate: { de: 'Startzeit', en: 'Start time' },
          endDate: { de: 'Endzeit', en: 'End time' },
          outlookBody: { de: 'Termin-Text', en: 'Calendar body' },
          location: { de: 'Ort', en: 'Location' },
          subject: { de: 'Betreff', en: 'Subject' },
          layout: { de: 'Kopfbild (Größe/Abstand)', en: 'Header image (size/spacing)' },
          organizer: { de: 'Organizer (im Termin-Text)', en: 'Organizer (in calendar body)' },
          logo: { de: 'Kopfbild', en: 'Header image' },
        };
        return (
          <Modal
            open={outlookConfirmOpen}
            onClose={cancelOutlookSave}
            maxWidth={620}
            dismissable={false}
            hideClose
            ariaLabel="Outlook-Update bestätigen"
            icon={<Calendar size={20} strokeWidth={2} />}
            title={<span id="outlook-confirm-title">{isDe ? 'Outlook-Termin der Teilnehmer aktualisieren?' : 'Update Outlook invite for attendees?'}</span>}
            subtitle={isDe
              ? 'Du hast Felder geändert, die im Outlook-Termin der Teilnehmer stehen. Hake an, welche Termine jetzt neu rausgehen — alles andere wird gespeichert, Outlook bleibt dort unangetastet. Nachholen geht jederzeit.'
              : 'You changed fields that appear in the attendees’ Outlook invites. Tick the invites to resend now — everything else is saved, Outlook is left alone there. You can resend later at any time.'}
            footer={<>
              <button type="button" className="btn btn-secondary" onClick={cancelOutlookSave}>
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button type="button" className="btn btn-primary" onClick={() => confirmOutlookSave()}>
                {checkedCount > 0
                  ? (isDe ? `Speichern & ${checkedCount} Termin${checkedCount === 1 ? '' : 'e'} aktualisieren` : `Save & update ${checkedCount} invite${checkedCount === 1 ? '' : 's'}`)
                  : (isDe ? 'Nur speichern' : 'Save only')}
              </button>
            </>}
          >
            {outlookConfirmOpen && (
              <div className="dex-ui-stack">
                {outlookConfirmItems.map(it => {
                  const changedLabels = it.changedFields.map(f => isDe ? fieldLabelMap[f].de : fieldLabelMap[f].en).join(', ');
                  const checked = !!outlookConfirmChecks[it.eventId];
                  // v15.3: leere changedFields-Liste = Item kommt aus dem
                  // persistierten OutlookDirty-Flag (frühere Session,
                  // wurde damals nicht synchronisiert). Klartext-Hinweis
                  // statt leerer „Geändert:"-Zeile.
                  const isFromPersistedDirty = !it.noOutlookYet && it.changedFields.length === 0;
                  // v11.69: noOutlookYet-Items bekommen wieder eine Checkbox.
                  // Default UNCHECKED. Beim Anhaken wird das Sub-Event in der
                  // Eventverwaltung komplett neu angelegt (DEX_Events-Item
                  // delete + create mit `existingSubsiteUrl`), damit der
                  // Outlook-Termin entsteht. Die bestehende Teilnehmerliste
                  // mit allen Anmeldungen bleibt unangetastet.
                  return (
                    <label key={it.eventId} className={cx('dex-ui-toggle-row', checked && 'is-active')}>
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={e => {
                          const next = e.target.checked;
                          setOutlookConfirmChecks(prev => ({ ...prev, [it.eventId]: next }));
                        }}
                      />
                      <span className="dex-ui-toggle-row-body">
                        <span className="dex-ui-toggle-row-title" style={{ wordBreak: 'break-word' }}>
                          {it.kind === 'top' && !it.noOutlookYet
                            ? (isDe ? `Hauptevent: ${it.title}` : `Main event: ${it.title}`)
                            : (isDe ? `Sub-Event: ${it.title}` : `Sub-event: ${it.title}`)}
                          {it.noOutlookYet && <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? 'noch kein Outlook-Termin' : 'no Outlook invite yet'}</span>}
                          {isFromPersistedDirty && <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? 'nicht synchronisiert' : 'not synced'}</span>}
                        </span>
                        {isFromPersistedDirty ? (
                          <span className="dex-ui-toggle-row-desc" style={{ display: 'block' }}>
                            {isDe
                              ? <><strong>Frühere Änderung nicht synchronisiert</strong> — beim letzten Speichern dieses Events wurden Outlook-relevante Felder geändert, der Outlook-Sync wurde aber damals übersprungen. Haken setzen, um die Teilnehmer jetzt nachträglich per Outlook-Update zu informieren.</>
                              : <><strong>Earlier change not yet synced</strong> — Outlook-relevant fields were changed in a previous save of this event, but the Outlook sync was skipped at the time. Tick the box to send the catch-up Outlook update to attendees now.</>}
                          </span>
                        ) : (
                          <span className="dex-ui-toggle-row-desc" style={{ display: 'block' }}>
                            {isDe ? 'Geändert: ' : 'Changed: '}{changedLabels}
                          </span>
                        )}
                        {it.noOutlookYet && (
                          <span className="dex-ui-callout dex-ui-callout--warn" style={{ marginTop: 8, fontSize: '0.76rem', padding: '8px 10px' }}>
                            {isDe
                              ? <>Für dieses Sub-Event gibt es noch keinen Outlook-Termin. Mit dem Haken wird es in der Eventverwaltung neu angelegt, damit der Termin entsteht. <strong>Die Teilnehmerliste mit allen Anmeldungen bleibt erhalten</strong> — nur die DEX_Events-Zeile bekommt eine neue ID.</>
                              : <>This sub-event has no Outlook event yet. Ticking re-creates it in the event admin so the invite can be generated. <strong>The participant list with all registrations stays intact</strong> — only the DEX_Events row gets a new ID.</>}
                          </span>
                        )}
                      </span>
                    </label>
                  );
                })}
                <div className="dex-ui-callout dex-ui-callout--neutral">
                  <span className="dex-ui-callout-icon"><Info size={16} /></span>
                  <span>
                    {isDe
                      ? <>Angehakt: Die Teilnehmer bekommen von Outlook eine Benachrichtigung &bdquo;Aktualisierter Termin&ldquo;. Nicht angehakt: Der Termin wird als &bdquo;ausstehender Outlook-Sync&ldquo; markiert — du kannst ihn später nachschicken.</>
                      : <>Ticked: attendees get an &ldquo;updated meeting&rdquo; notification from Outlook. Unticked: the invite is flagged as &ldquo;pending Outlook sync&rdquo; so you can resend it later.</>}
                  </span>
                </div>
              </div>
            )}
          </Modal>
        );
      })()}

      {/* v17.3: Unsaved-Changes-Confirm-Modal. Erscheint, wenn der User
          auf „Zurück" klickt und das Formular gegenüber dem Initial-
          Snapshot Änderungen hat.
          v30.1: Drei modusabhängige Wege, gestapelt statt nebeneinander:
           - Neu-Anlage: „Entwurf speichern" legt den Stand in den
             Entwurfs-Zwischenspeicher (v30.0) und verlässt den Wizard —
             beim nächsten Öffnen der Event-Erstellung wird er angeboten.
             „Event verwerfen" löscht den Entwurf endgültig.
           - Edit: „Änderungen speichern" = attemptSubmit wie bisher
             (blockt die Back-Nav, nach erfolgreichem Save navigiert der
             submit-success-Dispatch); „Änderungen verwerfen" verlässt
             ohne Speichern. */}
      {unsavedConfirmOpen && (
        <Modal
          open={true}
          onClose={() => { unsavedConfirmOpen.resolve(false); setUnsavedConfirmOpen(null); }}
          maxWidth={540}
          padding={24}
          ariaLabel={isDe ? 'Ungespeicherte Änderungen' : 'Unsaved changes'}
          icon={<AlertCircle size={20} />}
          title={isDe
            ? (isEditMode ? 'Ungespeicherte Änderungen' : 'Entwurf noch nicht gespeichert')
            : (isEditMode ? 'Unsaved changes' : 'Draft not saved yet')}
          subtitle={isDe
            ? (isEditMode ? 'Du hast Änderungen am Event vorgenommen, die noch nicht gespeichert sind.' : 'Dein Event ist noch nicht angelegt.')
            : (isEditMode ? 'You have made changes to this event that are not saved yet.' : 'Your event is not created yet.')}
          // v31.2: Drei Wege im Fuß statt gestapelt — Verwerfen steht links
          // mit Abstand, damit es nie direkt neben „Fortsetzen" liegt.
          footer={<>
            <div className="dex-ui-modal-foot-left">
              <button
                type="button"
                className="btn btn-danger"
                onClick={() => {
                  if (!isEditMode) {
                    // Verwerfen heisst verwerfen — auch den Entwurfs-
                    // Zwischenspeicher, sonst bietet ihn der naechste
                    // Besuch wieder an.
                    try { localStorage.removeItem(DRAFT_KEY); } catch { /* */ }
                  }
                  unsavedConfirmOpen.resolve(true);
                  setUnsavedConfirmOpen(null);
                }}
              >
                {isDe
                  ? (isEditMode ? 'Änderungen verwerfen' : 'Event verwerfen')
                  : (isEditMode ? 'Discard changes' : 'Discard event')}
              </button>
            </div>
            <button
              type="button"
              className="btn btn-secondary"
              onClick={() => { unsavedConfirmOpen.resolve(false); setUnsavedConfirmOpen(null); }}
            >
              {isDe
                ? (isEditMode ? 'Weiter bearbeiten' : 'Weiter erstellen')
                : (isEditMode ? 'Continue editing' : 'Continue creating')}
            </button>
            {isEditMode ? (
              /* v17.7: blockt die laufende Back-Nav (resolve(false)) und
                 triggert attemptSubmit; nach erfolgreichem Save dispatched
                 EventCreationPage „dex-event-submit-success" und
                 DexEventPlatform navigiert zum Organizer-Menü. */
              <button
                type="button"
                className="btn btn-primary"
                onClick={() => {
                  unsavedConfirmOpen.resolve(false);
                  setUnsavedConfirmOpen(null);
                  window.setTimeout(() => { attemptSubmit(); }, 0);
                }}
              >
                <Send size={14} /> {isDe ? 'Änderungen speichern' : 'Save changes'}
              </button>
            ) : (
              <button
                type="button"
                className="btn btn-primary"
                onClick={() => {
                  // Sofort schreiben — der 1,5-s-Debounce des Autosaves hat
                  // die letzten Eingaben sonst evtl. noch nicht gesichert.
                  try { localStorage.setItem(DRAFT_KEY, JSON.stringify({ savedAt: Date.now(), data: buildDraftPayload() })); } catch { /* best-effort */ }
                  unsavedConfirmOpen.resolve(true);
                  setUnsavedConfirmOpen(null);
                }}
              >
                <Send size={14} /> {isDe ? 'Entwurf speichern' : 'Save draft'}
              </button>
            )}
          </>}
        >
          <div className="dex-ui-callout dex-ui-callout--neutral">
            <span className="dex-ui-callout-icon"><Info size={16} /></span>
            <span>
              {isDe
                ? (isEditMode
                  ? <><strong>Änderungen speichern</strong> speichert und verlässt den Assistenten. <strong>Änderungen verwerfen</strong> verlässt ihn ohne zu speichern — das Event bleibt, wie es zuletzt gespeichert war.</>
                  : <><strong>Entwurf speichern</strong> behält den Stand — beim nächsten Öffnen der Event-Erstellung machst du genau hier weiter. Hochgeladene Bilder sind im Entwurf nicht enthalten. <strong>Event verwerfen</strong> löscht auch den Entwurf.</>)
                : (isEditMode
                  ? <><strong>Save changes</strong> saves and leaves the wizard. <strong>Discard changes</strong> leaves without saving — the event stays as it was last saved.</>
                  : <><strong>Save draft</strong> keeps this state — next time you open event creation you continue right here. Uploaded images are not part of the draft. <strong>Discard event</strong> also deletes the draft.</>)}
            </span>
          </div>
        </Modal>
      )}
    </>
  );
};
