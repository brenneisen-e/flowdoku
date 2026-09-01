/* WizardModals — aus EventCreationPage.tsx ausgelagert (Zeilen 6907-8577 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; die Komponente
 * gibt ein Fragment zurueck, damit die Geschwister-Reihenfolge im Elternbaum
 * exakt bleibt. */
import * as React from 'react';
import { Plus, Send, X } from '../Icons';
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
import { de } from 'date-fns/locale';

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
  autoDeregisterOnDecline: boolean;
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
  disableCancellationEmail: boolean;
  disableEmails: boolean;
  disableOutlook: boolean;
  disableRegistrationEmail: boolean;
  documents: { name: string; file?: File; url: string; size: number; }[];
  DRAFT_KEY: string;
  dragOverSectionId: string;
  dragSectionId: string;
  durchstarterCapacity: string;
  emailLanguage: string;
  emailLogoPreview: any;
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
  inactiveHandling: "notify" | "autoderegister";
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
  const { activeCommTabIdx, activeFrom, addrCity, addrHouseNo, addrStreet, addrZip, addSelectedSuggestedFields, agenda, applySubTransfer, askSalutation, attemptSubmit, audience, autoDeregisterOnDecline, berlinLocalToUtcIso, bilingualFields, buildDraftPayload, bulkOrganizerOpen, bulkQrScannerOpen, bulkTestTeamOpen, cancelOutlookSave, childTermPlural, childTermSingular, closeVisCopy, confirmOutlookSave, contactEmail, customFields, DEMO_VARIANTS, description, disableCancellationEmail, disableEmails, disableOutlook, disableRegistrationEmail, documents, DRAFT_KEY, dragOverSectionId, dragSectionId, durchstarterCapacity, emailLanguage, emailLogoPreview, emailTemplateOverrides, emailTemplates, endDate, eventImageUrl, excludedUsers, filterMode, funstarterCapacity, headerImageLayout, htmlEditorMode, htmlEditorOpen, htmlEditorTemplateType, imagePreview, inactiveHandling, isDe, isEditMode, isFictive, isMobile, isoToLocal, lastDeregisterDate, location, locationFilter, maxParticipants, newSectionError, newSectionModalOpen, newSectionName, organizer, organizerEmails, outlookBody, outlookConfirmChecks, outlookConfirmItems, outlookConfirmOpen, outlookEndOverride, outlookHeading, outlookLocationOverride, outlookLogoPreview, outlookStartOverride, outlookSubheading, outlookSubject, pendingSections, pendingSuccessDispatch, pendingSuccessDispatchRef, previewSections, qrScannerEmails, qrScannerNames, quiz, registrationDeadline, registrationLanguage, renderPreviewSection, requireSubEventSelection, scDescription, scopeSub, searchUsers, setBulkOrganizerOpen, setBulkQrScannerOpen, setBulkTestTeamOpen, setDragOverSectionId, setDragSectionId, setEmailTemplateOverrides, setHeaderImageLayout, setHtmlEditorOpen, setNewSectionError, setNewSectionModalOpen, setNewSectionName, setOrganizer, setOrganizerEmails, setOutlookBody, setOutlookConfirmChecks, setOutlookEndOverride, setOutlookHeading, setOutlookLocationOverride, setOutlookStartOverride, setOutlookSubheading, setOutlookSubject, setPendingSections, setPendingSuccessDispatch, setPreviewSections, setQrScannerEmails, setQrScannerNames, setScDescription, setShowB2runSuggested, setShowConfigCheck, setShowDemoVariantModal, setShowPreview, setShowRegisterPreview, setShowSuggestedModal, setShowSummaryModal, setSubEvents, setSubTransfer, setSuggestedSelection, setTestTeamEmails, setTestTeamNames, setUnsavedConfirmOpen, showB2runSuggested, showConfigCheck, showDemoVariantModal, showPreview, showRegisterPreview, showSuggestedModal, showSummaryModal, splitLabelA, splitLabelB, splitSharedWaitlist, startDate, SUB_TRANSFER_GROUPS, subEvents, subEventsOnlyMode, subGroupDiffCount, subTransfer, SUGGESTED_FIELDS_CATALOG, suggestedSelection, t, teamRegistrationEnabled, teamSize, testTeamEmails, testTeamNames, title, transferTimes, unlimitedParticipants, unsavedConfirmOpen, useSplitCapacities, visCopyModalOpen, waitlistEnabled, allowAttendeeUpload, askTeamName, attendeeUploadHint, attendeeUploadLabel, contactInfo, contactName, notifyOrgCancelMode, notifyOrgRegisterFromDate, notifyOrgRegisterMode, quizClusterSize, splitDescA, splitDescB, splitDisplayOrderReversed, splitHelpText, splitSectionTitle, teamJoinRequiresApproval, teamOpenSlotsVisible, teamPartialAllowed } = p;
  return (
    <>
      {/* ===== Vollbild-Vorschau Modal ===== */}
      {showPreview && (
        <div className="preview-modal" style={{
          position: 'fixed', inset: 0, background: '#fff', zIndex: 1000,
          display: 'flex', flexDirection: 'column',
        }}>
          <div className="preview-modal-inner" style={{
            background: '#fff', borderRadius: 0, width: '100%', maxWidth: '100%',
            height: '100%', overflow: 'auto', padding: 0,
          }}>
            <div style={{
              padding: '16px 24px', borderBottom: '1px solid var(--dex-gray-200)',
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
              position: 'sticky', top: 0, background: '#fff', zIndex: 1, borderRadius: '16px 16px 0 0',
            }}>
              <div>
                <h3 style={{ margin: 0 }}>Vorschau: Registrierungsseite</h3>
                <p style={{ margin: '4px 0 0', fontSize: '0.8rem', color: 'var(--dex-gray-400)' }}>
                  Sektionen per Drag &amp; Drop verschieben
                </p>
              </div>
              <button onClick={() => setShowPreview(false)} style={{ background: 'none', border: 'none', fontSize: '1.5rem', cursor: 'pointer', color: 'var(--dex-gray-500)' }}>
                <X size={24} />
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
                    borderTop: dragOverSectionId === section.id ? '3px solid var(--dex-green)' : undefined,
                    cursor: 'grab',
                    position: 'relative',
                  }}
                >
                  <div style={{
                    position: 'absolute', top: 4, right: 8, fontSize: '0.65rem',
                    color: 'var(--dex-gray-300)', fontWeight: 600, userSelect: 'none',
                  }}>
                    ⠿ verschieben
                  </div>
                  {renderPreviewSection(section.id)}
                </div>
              ))}
            </div>

            <div style={{
              padding: '16px 24px', borderTop: '1px solid var(--dex-gray-200)',
              display: 'flex', gap: 12, justifyContent: 'flex-end',
              position: 'sticky', bottom: 0, background: '#fff', borderRadius: '0 0 16px 16px',
            }}>
              <button className="btn btn-secondary" onClick={() => setShowPreview(false)}>
                Zurück zum Formular
              </button>
              <button
                className="btn btn-primary"
                disabled={!title}
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
          showTimeSelect: true, timeFormat: 'HH:mm', timeIntervals: 15, timeCaption: 'Uhrzeit',
          dateFormat: 'dd.MM.yyyy, HH:mm', locale: 'de', className: 'form-input',
          wrapperClassName: 'dex-datepicker-wrapper', calendarClassName: 'dex-datepicker-calendar',
          popperPlacement: 'bottom-start' as const, isClearable: true, autoComplete: 'off',
        };
        const outlookDateEditor = (
          <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
              <label style={{ fontSize: '0.68rem', color: 'var(--dex-gray-400)' }}>Start</label>
              <DatePicker {...dpCommon} selected={olIsoToDate(olStartOverrideVal)} onChange={(d: Date | null) => setOlStart(olDateToIso(d))} placeholderText={olStart ? olFmt(olStart) + ' (übernommen)' : 'Start'} />
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
              <label style={{ fontSize: '0.68rem', color: 'var(--dex-gray-400)' }}>Ende</label>
              <DatePicker {...dpCommon} selected={olIsoToDate(olEndOverrideVal)} onChange={(d: Date | null) => setOlEnd(olDateToIso(d))} placeholderText={olEnd ? olFmt(olEnd) + ' (übernommen)' : 'Ende'} />
            </div>
          </div>
        );
        // v18.46: Standard-Body-Vorlage (mit Platzhaltern) für „Standardtext laden"
        // im Outlook-Editor — Sprache folgt der aktiven Mail-Sprache.
        const outlookDefaultBody = (emailLanguage === 'EN')
          ? '<p>You are registered for the event <strong>{{EventTitle}}</strong>.</p>'
            + '<p>If you are unable to attend, please cancel your registration in time via the <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView" style="color:#86bc25;font-weight:600;">DEX App</a> („My Events").</p>'
            + '<p>For organizational questions please contact <strong>{{Organizer}}</strong>.</p>'
          : '<p>Ihr seid für das Event <strong>{{EventTitle}}</strong> angemeldet.</p>'
            + '<p>Falls ihr nicht teilnehmen könnt, meldet euch bitte rechtzeitig über die <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView" style="color:#86bc25;font-weight:600;">DEX App</a> („Meine Events") ab.</p>'
            + '<p>Bei organisatorischen Fragen wendet euch bitte an <strong>{{Organizer}}</strong>.</p>';
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
              ? 'Outlook-Termin: Body bearbeiten'
              : isDescription
                ? (scopeSub
                  ? (isDe
                    ? `Beschreibung: ${shortSubEventTitle(scopeSub.title, title) || (childTermSingular || 'Sub-Event')}`
                    : `Description: ${shortSubEventTitle(scopeSub.title, title) || (childTermSingular || 'sub-event')}`)
                  : (isDe ? 'Event-Beschreibung bearbeiten' : 'Edit event description'))
                : `E-Mail-Template: ${tType}`}
            // v28.7: Die Starthilfe (Tipp-Text + Vorschlags-Chips) lebt jetzt
            // HIER im Editor statt als Dauer-Box im Wizard-Schritt.
            headerExtra={isDescription ? (
              <div style={{ padding: '10px 12px', borderRadius: 'var(--dex-radius)', background: 'var(--dex-gray-50, #f7f7f7)', border: '1px solid var(--dex-gray-200)', fontSize: '0.78rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                {isDe
                  ? <>Die Beschreibung ist der <strong>einladende Einleitungstext ganz oben auf der Anmeldemaske</strong> — das Erste, was deine Teilnehmenden lesen. Erzähl hier gern, <strong>worum es geht, für wen das Event ist und was man wissen sollte</strong>.<br />Ein kleiner Tipp: <strong>Zeitpunkt, Ort, Organizer und Kontaktperson musst du hier nicht angeben</strong> — die zeigt die App bereits als eigene Felder darüber an. So bleibt dein Text schön schlank und einladend.</>
                  : <>The description is the <strong>inviting intro text right at the top of the registration form</strong> — the first thing your attendees read. Feel free to tell them <strong>what the event is about, who it&rsquo;s for and what to know</strong>.<br />A little tip: <strong>you don&rsquo;t need to add the date, location, organizer or contact person here</strong> — the app already shows those as their own fields above. That keeps your text nice and inviting.</>}
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
            ] : [
              { key: '{{Name}}', label: 'Name' },
              { key: '{{EventTitle}}', label: 'Event' },
              { key: '{{Organizer}}', label: 'Organizer' },
              ...(contactEmail.trim() ? [{ key: '{{ContactEmail}}', label: 'Kontakt-Mail' }] : []),
              { key: '{{AppUrl}}', label: 'App Link' },
              { key: '{{WaitlistPosition}}', label: 'Waitlist #' },
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
        title="Massenimport — Co-Organizer"
        description={(
          <p style={{ marginTop: 0 }}>
            Mehrere <strong>Co-Organizer</strong> auf einmal hinzufügen. Reihenfolge
            spielt eine Rolle — der erste Eintrag in der Liste bleibt der Haupt-Organizer.
            Massenimport hängt neue Personen <strong>hinten</strong> an.
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
        title="Massenimport — Test-Team"
        description={(
          <p style={{ marginTop: 0 }}>
            Mehrere <strong>Test-Team-Mitglieder</strong> auf einmal hinzufügen. Test-Team
            sieht das Event schon im Entwurfsmodus und kann sich testweise anmelden.
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
        title="Massenimport — Check-In Team"
        description={(
          <p style={{ marginTop: 0 }}>
            Mehrere <strong>Check-In-Team-Mitglieder</strong> auf einmal hinzufügen. Diese
            Personen dürfen am Eventtag den QR-Scanner / Check-In-Tool benutzen, haben aber
            keine weiteren Admin-Rechte.
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
        sections.push({
          title: isDe ? 'Schritt 2 — Organizer & Team' : 'Step 2 — Organizers & Team',
          rows: [
            { label: 'Organizer', value: orgList.length ? `${orgList.length} ${isDe ? 'Person(en)' : 'person(s)'}` : '—', status: orgList.length ? 'ok' : 'missing' },
            { label: isDe ? 'Test-Team' : 'Test team', value: testTeamEmails.length ? `${testTeamEmails.length} ${isDe ? 'Person(en)' : 'person(s)'}` : '—', status: testTeamEmails.length ? 'ok' : 'empty' },
            { label: isDe ? 'Check-In-Team' : 'Check-in team', value: qrScannerEmails.length ? `${qrScannerEmails.length} ${isDe ? 'Person(en)' : 'person(s)'}` : '—', status: qrScannerEmails.length ? 'ok' : 'empty' },
          ],
        });
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
            { label: isDe ? 'Mail-Sprache' : 'Email language', value: (emailLanguage || 'EN').toUpperCase() === 'DE' ? 'Deutsch' : 'English', status: 'ok' },
            { label: isDe ? 'Bestätigungs-Mails' : 'Confirmation emails', value: disableEmails ? (isDe ? 'deaktiviert' : 'disabled') : (isDe ? 'aktiv' : 'on'), status: disableEmails ? 'ok' : 'default' },
            ...(!disableEmails && (disableRegistrationEmail || disableCancellationEmail) ? [{ label: isDe ? 'Einzeln deaktiviert' : 'Individually disabled', value: [disableRegistrationEmail ? (isDe ? 'Anmelde-Bestätigung' : 'registration confirmation') : '', disableCancellationEmail ? (isDe ? 'Abmelde-Bestätigung' : 'cancellation confirmation') : ''].filter(Boolean).join(', '), status: 'ok' as CheckStatus }] : []),
            { label: isDe ? 'Outlook-Termin' : 'Outlook invite', value: disableOutlook ? (isDe ? 'deaktiviert' : 'disabled') : (isDe ? 'aktiv' : 'on'), status: disableOutlook ? 'ok' : 'default' },
            { label: isDe ? 'Auto-Abmeldung bei Outlook-Absage' : 'Auto-cancel on Outlook decline', value: autoDeregisterOnDecline ? (isDe ? 'an' : 'on') : (isDe ? 'aus' : 'off'), status: autoDeregisterOnDecline ? 'ok' : 'default' },
            { label: isDe ? 'Person nicht mehr bei Deloitte' : 'Person no longer at Deloitte', value: inactiveHandling === 'autoderegister' ? (isDe ? 'automatisch abmelden' : 'auto-deregister') : (isDe ? 'Organizer informieren' : 'notify organizer'), status: inactiveHandling === 'autoderegister' ? 'ok' : 'default' },
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
        const allRows = sections.reduce((acc, s) => acc + s.rows.length, 0);
        void allRows;
        const missingCount = sections.reduce((acc, s) => acc + s.rows.filter(r => r.status === 'missing').length, 0);
        const emptyCount = sections.reduce((acc, s) => acc + s.rows.filter(r => r.status === 'empty').length, 0);
        const chip = (st: CheckStatus): React.ReactElement | null => {
          if (st === 'default') return <span style={{ fontSize: '0.66rem', fontWeight: 700, padding: '1px 8px', borderRadius: 999, background: 'var(--dex-gray-100)', color: 'var(--dex-gray-500)', flexShrink: 0 }}>{isDe ? 'Standard' : 'Default'}</span>;
          if (st === 'empty') return <span style={{ fontSize: '0.66rem', fontWeight: 700, padding: '1px 8px', borderRadius: 999, background: 'rgba(237,139,0,0.12)', color: 'var(--dex-orange-dark, #b35a00)', flexShrink: 0 }}>{isDe ? 'leer (optional)' : 'empty (optional)'}</span>;
          if (st === 'missing') return <span style={{ fontSize: '0.66rem', fontWeight: 700, padding: '1px 8px', borderRadius: 999, background: 'rgba(218,41,28,0.12)', color: 'var(--dex-red, #c00)', flexShrink: 0 }}>{isDe ? 'fehlt' : 'missing'}</span>;
          return null;
        };
        return (
          <div
            role="dialog"
            aria-modal="true"
            onClick={() => setShowConfigCheck(false)}
            style={{
              position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1250,
              display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
            }}
          >
            <div
              onClick={e => e.stopPropagation()}
              className="card"
              style={{ width: '100%', maxWidth: 760, maxHeight: '88vh', overflow: 'auto', padding: 24, borderRadius: 16, background: '#fff', boxShadow: '0 20px 60px rgba(0,0,0,0.25)' }}
            >
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 4 }}>
                <h3 style={{ margin: 0, fontSize: '1.1rem' }}>{isDe ? 'Event prüfen — alle Einstellungen im Überblick' : 'Review event — all settings at a glance'}</h3>
                <button type="button" onClick={() => setShowConfigCheck(false)} style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 4 }} aria-label={isDe ? 'Schließen' : 'Close'}><X size={20} /></button>
              </div>
              <p style={{ margin: '0 0 14px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {missingCount > 0
                  ? (isDe ? <><strong style={{ color: 'var(--dex-red, #c00)' }}>{missingCount} Pflichtangabe(n) fehlen</strong>{emptyCount > 0 ? <> · {emptyCount} optionale Punkte sind noch leer</> : null}.</> : <><strong style={{ color: 'var(--dex-red, #c00)' }}>{missingCount} required item(s) missing</strong>{emptyCount > 0 ? <> · {emptyCount} optional items still empty</> : null}.</>)
                  : emptyCount > 0
                    ? (isDe ? <>Alle Pflichtangaben sind gesetzt — <strong>{emptyCount} optionale Punkte</strong> sind noch leer.</> : <>All required items are set — <strong>{emptyCount} optional items</strong> are still empty.</>)
                    : (isDe ? 'Alles gesetzt — keine offenen Punkte.' : 'Everything set — nothing open.')}
              </p>
              {sections.map((sec, si) => (
                <div key={si} style={{ marginBottom: 14 }}>
                  <div style={{ fontWeight: 700, fontSize: '0.85rem', color: 'var(--dex-green-dark, #4a7c1f)', borderBottom: '1px solid var(--dex-gray-100)', paddingBottom: 4, marginBottom: 6 }}>{sec.title}</div>
                  {sec.rows.map((r, ri) => (
                    <div key={ri} style={{ display: 'flex', flexDirection: isMobile ? 'column' : 'row', alignItems: 'flex-start', gap: isMobile ? 2 : 10, padding: '4px 0', fontSize: '0.8rem' }}>
                      <span style={{ flex: isMobile ? '0 0 auto' : '0 0 230px', width: isMobile ? '100%' : undefined, color: 'var(--dex-gray-500)' }}>{r.label}</span>
                      <span style={{ flex: 1, color: 'var(--dex-gray-800)', minWidth: 0, overflowWrap: 'anywhere' }}>{r.value}</span>
                      {chip(r.status)}
                    </div>
                  ))}
                </div>
              ))}
              <div style={{ textAlign: 'right', marginTop: 6 }}>
                <button className="btn btn-primary" onClick={() => setShowConfigCheck(false)}>{isDe ? 'Schließen' : 'Close'}</button>
              </div>
            </div>
          </div>
        );
      })()}

      {showDemoVariantModal && (
        <div
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1200,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
          }}
          onClick={() => setShowDemoVariantModal(false)}
        >
          <div
            className="card"
            style={{
              width: '100%', maxWidth: 760, maxHeight: '90vh', overflow: 'auto',
              padding: 24, borderRadius: 16, background: '#fff',
              boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
            }}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0, color: 'var(--dex-green-dark, #4a7c1f)' }}>
                {isDe ? 'Demo-Daten laden' : 'Load demo data'}
              </h3>
              <button
                style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-600)' }}
                onClick={() => setShowDemoVariantModal(false)}
                aria-label={isDe ? 'Schließen' : 'Close'}
              >
                <X size={20} />
              </button>
            </div>
            <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
              {isDe
                ? 'Wähle eine Vorlage. Die ausgewählte Variante überschreibt deine aktuellen Eingaben — verworfen wird nichts, falls du noch nichts gespeichert hast.'
                : 'Choose a template. The selected variant overrides your current input — nothing is lost if you haven\'t saved yet.'}
            </p>
            {(() => {
              const cards: Array<{ key: keyof typeof DEMO_VARIANTS; titleDe: string; titleEn: string; descDe: string; descEn: string }> = [
                {
                  key: 'standard',
                  titleDe: 'Standard',
                  titleEn: 'Standard',
                  descDe: 'Ein Event, eine Gruppe. Typisches Meeting / Lunch.',
                  descEn: 'One event, one group. Typical meeting or lunch.',
                },
                {
                  key: 'groups',
                  titleDe: 'Mit Gruppen',
                  titleEn: 'With groups',
                  descDe: 'Event mit zwei Teilnehmer-Gruppen (Split Capacity), z.B. Vormittag / Nachmittag.',
                  descEn: 'Event with two participant groups (split capacity), e.g. morning / afternoon.',
                },
                {
                  key: 'subevent',
                  titleDe: 'Mit Sub-Event',
                  titleEn: 'With sub-event',
                  descDe: 'Haupt-Event + 1 Sub-Event, z.B. Conference + Dinner.',
                  descEn: 'Main event + 1 sub-event, e.g. conference + dinner.',
                },
                {
                  key: 'subeventTeam',
                  titleDe: 'Mit Sub-Event + Team',
                  titleEn: 'With sub-event + team',
                  descDe: 'Wie links, aber mit Team-Anmeldung (Teams à 4 Personen).',
                  descEn: 'Same as on the left, but with team registration (teams of 4 people).',
                },
              ];
              return (
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: 12 }}>
                  {cards.map(card => (
                    <button
                      key={card.key}
                      type="button"
                      onClick={() => {
                        DEMO_VARIANTS[card.key]();
                        setShowDemoVariantModal(false);
                      }}
                      style={{
                        textAlign: 'left',
                        padding: 16,
                        borderRadius: 12,
                        border: '1px solid var(--dex-gray-200)',
                        background: '#fff',
                        cursor: 'pointer',
                        transition: 'all 0.15s ease',
                        display: 'flex',
                        flexDirection: 'column',
                        gap: 8,
                        minHeight: 120,
                      }}
                      onMouseEnter={e => {
                        e.currentTarget.style.borderColor = 'var(--dex-green-dark, #4a7c1f)';
                        e.currentTarget.style.boxShadow = '0 6px 16px rgba(74,124,31,0.12)';
                        e.currentTarget.style.transform = 'translateY(-2px)';
                      }}
                      onMouseLeave={e => {
                        e.currentTarget.style.borderColor = 'var(--dex-gray-200)';
                        e.currentTarget.style.boxShadow = 'none';
                        e.currentTarget.style.transform = 'translateY(0)';
                      }}
                    >
                      <div style={{ fontWeight: 700, fontSize: '0.95rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
                        {isDe ? card.titleDe : card.titleEn}
                      </div>
                      <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                        {isDe ? card.descDe : card.descEn}
                      </div>
                    </button>
                  ))}
                </div>
              );
            })()}
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 20, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
              <button
                type="button"
                className="btn btn-outline"
                onClick={() => setShowDemoVariantModal(false)}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
            </div>
          </div>
        </div>
      )}

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
        return (
          <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 1300, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}>
            <div className="card" style={{ width: '100%', maxWidth: 680, maxHeight: '88vh', overflow: 'auto', padding: 24, borderRadius: 14, background: '#fff' }}>
              <h3 style={{ margin: '0 0 4px', fontSize: '1.1rem' }}>
                {isDe ? 'Einstellungen übertragen' : 'Transfer settings'}
              </h3>
              <p style={{ margin: '0 0 16px', fontSize: '0.84rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {isDe
                  ? <>Die ausgewählten Einstellungen von <strong>„{srcName}“</strong> werden auf die ausgewählten {childTermPlural || 'Sub-Events'} übertragen und <strong>überschreiben</strong> die dortigen Werte. Gespeichert wird erst, wenn du den Assistenten speicherst.</>
                  : <>The selected settings of <strong>„{srcName}“</strong> are applied to the selected sub-events and <strong>overwrite</strong> their values. Nothing is stored until you save the wizard.</>}
              </p>

              <div style={{ fontWeight: 700, fontSize: '0.85rem', marginBottom: 8 }}>
                {isDe ? '1. Was soll übertragen werden?' : '1. What should be transferred?'}
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 18 }}>
                {SUB_TRANSFER_GROUPS.map(g => {
                  const on = subTransfer.groups.indexOf(g.key) >= 0;
                  const n = subGroupDiffCount(subTransfer.fromIdx, g.fields);
                  return (
                    <label key={g.key} style={{
                      display: 'flex', alignItems: 'flex-start', gap: 10, padding: '8px 12px', borderRadius: 8,
                      border: `1px solid ${on ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                      background: on ? 'rgba(134,188,37,0.06)' : '#fff', cursor: 'pointer', fontSize: '0.84rem',
                    }}>
                      <input type="checkbox" checked={on} onChange={() => toggleGroup(g.key)} style={{ marginTop: 3 }} />
                      <span>
                        {isDe ? g.de : g.en}
                        {n > 0 && (
                          <span style={{ marginLeft: 8, fontSize: '0.72rem', fontWeight: 700, padding: '1px 7px', borderRadius: 10, background: '#fff3d6', color: '#7a5a12', border: '1px solid #e0b34d' }}>
                            {isDe ? `${n}× abweichend` : `${n}× differing`}
                          </span>
                        )}
                      </span>
                    </label>
                  );
                })}
              </div>

              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
                <div style={{ fontWeight: 700, fontSize: '0.85rem' }}>
                  {isDe ? '2. Auf welche übertragen?' : '2. Transfer to which ones?'}
                </div>
                <button
                  type="button"
                  className="btn btn-secondary"
                  style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                  onClick={() => setSubTransfer(prev => prev && ({
                    ...prev,
                    targets: prev.targets.length === subEvents.length - 1
                      ? []
                      : subEvents.map((_, i) => i).filter(i => i !== prev.fromIdx),
                  }))}
                >
                  {subTransfer.targets.length === subEvents.length - 1
                    ? (isDe ? 'Keine' : 'None')
                    : (isDe ? 'Alle' : 'All')}
                </button>
              </div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 20 }}>
                {subEvents.map((s, i) => {
                  if (i === subTransfer.fromIdx) return null;
                  const on = subTransfer.targets.indexOf(i) >= 0;
                  const nm = shortSubEventTitle(s.title, title) || (isDe ? 'Ohne Titel' : 'Untitled');
                  return (
                    <label key={s.id || i} style={{
                      display: 'inline-flex', alignItems: 'center', gap: 6, padding: '6px 10px', borderRadius: 8,
                      border: `1px solid ${on ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                      background: on ? 'rgba(134,188,37,0.10)' : '#fff', cursor: 'pointer', fontSize: '0.8rem',
                    }}>
                      <input type="checkbox" checked={on} onChange={() => toggleTarget(i)} />
                      {nm}
                    </label>
                  );
                })}
              </div>

              <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                <button type="button" className="btn btn-secondary" onClick={() => setSubTransfer(null)}>
                  {isDe ? 'Abbrechen' : 'Cancel'}
                </button>
                <button type="button" className="btn btn-primary" disabled={!canApply} onClick={applySubTransfer}>
                  {isDe
                    ? `Auf ${subTransfer.targets.length} übertragen`
                    : `Transfer to ${subTransfer.targets.length}`}
                </button>
              </div>
            </div>
          </div>
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
          <div
            style={{
              position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1300,
              display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
            }}
            onClick={closeAndDispatch}
          >
            <div
              className="card"
              style={{
                width: '100%', maxWidth: 560, padding: 24, borderRadius: 16,
                background: '#fff', boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
              }}
              onClick={e => e.stopPropagation()}
            >
              <h3 style={{ margin: 0, color: 'var(--dex-green-dark, #4a7c1f)' }}>
                {isDe ? 'Event-Zusammenfassung herunterladen?' : 'Download event summary?'}
              </h3>
              <p style={{ marginTop: 12, color: 'var(--dex-gray-700)', lineHeight: 1.55, fontSize: '0.95rem' }}>
                {isDe
                  ? <>Das Event wurde gespeichert. Möchtest du jetzt eine <strong>A4-Zusammenfassung</strong> mit allen Sektionen (Foto, Beschreibung, Sichtbarkeit, Felder, Kommunikation, Dokumente, Sub-Events…) herunterladen? Du kannst sie z.B. einem Partner zur Durchsicht weiterleiten.</>
                  : <>The event has been saved. Would you like to download a <strong>one-page A4 summary</strong> with every section (photo, description, visibility, fields, communication, documents, sub-events…)? You can forward it to a partner for review.</>}
              </p>
              <div style={{
                marginTop: 18, padding: '10px 14px', background: 'rgba(0,90,156,0.06)',
                border: '1px solid rgba(0,90,156,0.25)', borderRadius: 8,
                fontSize: '0.82rem', color: 'var(--dex-gray-700)',
              }}>
                {isDe
                  ? <><strong>Hinweis:</strong> Beim PDF-Export öffnet sich der Browser-Druckdialog. Wähle dort <strong>&bdquo;Als PDF speichern&ldquo;</strong> als Ziel. Word-Export lädt direkt eine .doc-Datei herunter.</>
                  : <><strong>Note:</strong> The PDF export opens the browser print dialog — pick <strong>&ldquo;Save as PDF&rdquo;</strong> as the destination. Word export downloads a .doc file directly.</>}
              </div>
              <div style={{ marginTop: 22, display: 'flex', flexWrap: 'wrap', gap: 10, justifyContent: 'flex-end' }}>
                <button className="btn btn-outline" onClick={closeAndDispatch}>
                  {isDe ? 'Nein, danke' : 'No, thanks'}
                </button>
                <button className="btn btn-secondary" onClick={onDoc}>
                  {isDe ? 'Als Word (.doc)' : 'As Word (.doc)'}
                </button>
                <button className="btn btn-primary" onClick={onPdf}>
                  {isDe ? 'Als PDF' : 'As PDF'}
                </button>
              </div>
            </div>
          </div>
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
        ariaLabel="Neuen Quiz-Bereich anlegen"
      >
        {newSectionModalOpen && (
          <>
            <h3 style={{ marginTop: 0, marginBottom: 8, fontSize: '1.15rem' }}>
              Neuen Bereich anlegen
            </h3>
            <p style={{ margin: '0 0 14px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
              Bereiche bündeln Quiz-Fragen auf einer gemeinsamen Seite. Vergib einen
              kurzen, sprechenden Namen — z.B. <em>Orte</em>, <em>Geschichte</em> oder <em>Foto-Quiz</em>.
            </p>
            <input
              type="text"
              className="form-input"
              autoFocus
              value={newSectionName}
              placeholder='z.B. "Orte"'
              onChange={e => { setNewSectionName(e.target.value); setNewSectionError(''); }}
              onKeyDown={e => {
                if (e.key === 'Enter') {
                  const name = newSectionName.trim();
                  if (!name) { setNewSectionError('Bitte einen Namen eingeben.'); return; }
                  const existing = new Set<string>();
                  for (const q of quiz) if (q.section) existing.add(q.section);
                  for (const p of pendingSections) existing.add(p);
                  if (existing.has(name)) { setNewSectionError('Ein Bereich mit diesem Namen existiert bereits.'); return; }
                  setPendingSections([...pendingSections, name]);
                  setNewSectionModalOpen(false);
                } else if (e.key === 'Escape') {
                  setNewSectionModalOpen(false);
                }
              }}
              style={{ fontSize: '0.95rem', marginBottom: 8 }}
            />
            {newSectionError && (
              <div style={{ color: 'var(--dex-red, #c00)', fontSize: '0.78rem', marginBottom: 10 }}>{newSectionError}</div>
            )}
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 12 }}>
              <button
                type="button"
                className="btn btn-secondary"
                onClick={() => setNewSectionModalOpen(false)}
              >
                Abbrechen
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={() => {
                  const name = newSectionName.trim();
                  if (!name) { setNewSectionError('Bitte einen Namen eingeben.'); return; }
                  const existing = new Set<string>();
                  for (const q of quiz) if (q.section) existing.add(q.section);
                  for (const p of pendingSections) existing.add(p);
                  if (existing.has(name)) { setNewSectionError('Ein Bereich mit diesem Namen existiert bereits.'); return; }
                  setPendingSections([...pendingSections, name]);
                  setNewSectionModalOpen(false);
                }}
              >
                <Plus size={14} /> Bereich anlegen
              </button>
            </div>
          </>
        )}
      </Modal>

      {/* Modal: Vorgeschlagene Felder auswählen (Multi-Select) — v13.4 über <Modal>. */}
      <Modal
        open={showSuggestedModal}
        onClose={() => setShowSuggestedModal(false)}
        maxWidth={540}
        ariaLabel="Vorgeschlagene Felder auswählen"
      >
        {showSuggestedModal && (
          <div
            style={{
              display: 'flex', flexDirection: 'column', gap: 14,
            }}
          >
            <h2 style={{ margin: 0, fontSize: '1.25rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Vorgeschlagene Felder' : 'Suggested fields'}
            </h2>
            <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
              {isDe
                ? 'Wähle aus dem Katalog, welche Felder dem Event hinzugefügt werden sollen. Du kannst die Felder danach weiter anpassen.'
                : 'Pick the fields you want to add to the event. You can still tweak them afterwards.'}
            </p>
            {/* v10.21: Catalog gruppiert nach Kategorie. Allgemeine Felder
                immer ausgeklappt, B2Run-Felder default eingeklappt mit
                Toggle. Jeder Eintrag bekommt ein Badge mit der Kategorie. */}
            {(() => {
              const generalEntries = SUGGESTED_FIELDS_CATALOG.filter(s => s.category === 'general');
              const b2runEntries = SUGGESTED_FIELDS_CATALOG.filter(s => s.category === 'b2run');
              const renderEntry = (s: SuggestedEntry): React.ReactElement => (
                <label
                  key={s.key}
                  style={{
                    display: 'flex', gap: 10, alignItems: 'flex-start',
                    padding: '10px 12px', border: '1px solid var(--dex-gray-200)',
                    borderRadius: 8, cursor: 'pointer',
                    background: suggestedSelection[s.key] ? 'var(--dex-gray-50, #fafafa)' : '#fff',
                  }}
                >
                  <input
                    type="checkbox"
                    checked={!!suggestedSelection[s.key]}
                    onChange={e => setSuggestedSelection({ ...suggestedSelection, [s.key]: e.target.checked })}
                    style={{ marginTop: 3, flexShrink: 0 }}
                  />
                  {/* v10.23: passendes Fluent-UI-Icon links neben dem Label,
                      damit die Auswahl auf einen Blick visuell wiedererkennbar
                      ist. Farbe analog zur Kategorie (grün=Allgemein,
                      orange=B2Run). */}
                  <Icon
                    iconName={s.icon}
                    style={{
                      fontSize: 20, flexShrink: 0, marginTop: 2,
                      color: s.category === 'b2run' ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-green-dark, #4a7c1f)',
                    }}
                  />
                  <span style={{ flex: 1 }}>
                    <span style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                      <strong style={{ fontSize: '0.9rem', color: 'var(--dex-gray-800)' }}>{s.label}</strong>
                      <span style={{
                        fontSize: '0.65rem', fontWeight: 600,
                        padding: '2px 8px', borderRadius: 999,
                        textTransform: 'uppercase', letterSpacing: 0.5,
                        background: s.category === 'b2run' ? 'rgba(237,139,0,0.12)' : 'rgba(134,188,37,0.12)',
                        color: s.category === 'b2run' ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-green-dark, #4a7c1f)',
                      }}>
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
                    <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>{s.description}</div>
                  </span>
                </label>
              );
              return (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 14, marginTop: 4 }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                    {generalEntries.map(renderEntry)}
                  </div>
                  <div style={{ borderTop: '1px solid var(--dex-gray-200)', paddingTop: 14 }}>
                    <button
                      type="button"
                      onClick={() => setShowB2runSuggested(v => !v)}
                      style={{
                        display: 'flex', alignItems: 'center', gap: 8,
                        background: 'none', border: 'none', padding: 0,
                        fontSize: '0.85rem', fontWeight: 600, cursor: 'pointer',
                        color: 'var(--dex-gray-700)',
                      }}
                    >
                      <span style={{ display: 'inline-flex', transform: showB2runSuggested ? 'rotate(90deg)' : 'rotate(0deg)', transition: 'transform 0.15s' }}>▶</span>
                      {isDe ? 'B2Run-spezifische Felder' : 'B2Run-specific fields'}
                      <span style={{
                        fontSize: '0.65rem', fontWeight: 600,
                        padding: '2px 8px', borderRadius: 999,
                        background: 'rgba(237,139,0,0.12)',
                        color: 'var(--dex-orange-dark, #b35a00)',
                      }}>
                        B2Run · {b2runEntries.length}
                      </span>
                    </button>
                    {showB2runSuggested && (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginTop: 10 }}>
                        <p style={{ margin: 0, fontSize: '0.78rem', color: 'var(--dex-gray-500)', lineHeight: 1.45 }}>
                          {isDe
                            ? 'Diese Felder sind speziell für B2Run-Lauf-Events vorgesehen (Startblock, Altersklasse, Datenschutz-Checkbox mit b2run.de-Links etc.). Bei normalen Events brauchst du sie nicht.'
                            : 'These fields are intended for B2Run running events (start block, age group, B2Run-specific privacy checkbox etc.). Skip them for standard events.'}
                        </p>
                        {b2runEntries.map(renderEntry)}
                      </div>
                    )}
                  </div>
                </div>
              );
            })()}
            <div style={{ display: 'flex', gap: 8, justifyContent: 'space-between', marginTop: 6 }}>
              <div style={{ display: 'flex', gap: 8 }}>
                <button
                  type="button"
                  className="btn btn-outline"
                  style={{ fontSize: '0.8rem', padding: '4px 12px' }}
                  onClick={() => {
                    const all: Record<string, boolean> = {};
                    for (const s of SUGGESTED_FIELDS_CATALOG) all[s.key] = true;
                    setSuggestedSelection(all);
                  }}
                >
                  {isDe ? 'Alle' : 'All'}
                </button>
                <button
                  type="button"
                  className="btn btn-outline"
                  style={{ fontSize: '0.8rem', padding: '4px 12px' }}
                  onClick={() => setSuggestedSelection({})}
                >
                  {isDe ? 'Keine' : 'None'}
                </button>
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setShowSuggestedModal(false)}
                >
                  {isDe ? 'Abbrechen' : 'Cancel'}
                </button>
                <button
                  type="button"
                  className="btn btn-primary"
                  onClick={addSelectedSuggestedFields}
                  disabled={!Object.values(suggestedSelection).some(Boolean)}
                >
                  {isDe ? 'Hinzufügen' : 'Add'}
                </button>
              </div>
            </div>
          </div>
        )}
      </Modal>

      {/* v22.62: „Sichtbarkeit auf Sub-Events übernehmen?" — erscheint beim
          ersten „Weiter"/Speichern, sobald die Klammer eine Sichtbarkeit hat
          und Sub-Events existieren. */}
      <Modal
        open={visCopyModalOpen}
        onClose={() => closeVisCopy(false)}
        maxWidth={560}
        dismissable={false}
        ariaLabel={isDe ? 'Sichtbarkeit übernehmen' : 'Apply visibility'}
      >
        {visCopyModalOpen && (
          <div>
            <h2 style={{ margin: '0 0 10px', fontSize: '1.15rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)' }}>
              {isDe ? 'Sichtbarkeit auf alle Sub-Events übernehmen?' : 'Apply visibility to all sub-events?'}
            </h2>
            <p style={{ margin: '0 0 16px', fontSize: '0.9rem', color: 'var(--dex-gray-700)', lineHeight: 1.55 }}>
              {isDe
                ? <>Du hast für {subEventsOnlyMode ? 'die Klammer' : 'das Hauptevent'} eine Sichtbarkeit gesetzt. Sollen <strong>alle {subEvents.length} Sub-Events</strong> dieselbe Sichtbarkeit (Standortfilter + Mailverteiler + Verknüpfung) übernehmen?<br /><br />Das ist meist sinnvoll, damit jeder, der das Event sehen soll, auch die Sub-Events erreicht — der Zugang läuft ohnehin über die Sichtbarkeit des Gesamt-Events. Bereits gesetzte, abweichende Sub-Event-Sichtbarkeiten werden dabei <strong>überschrieben</strong>.</>
                : <>You set a visibility for {subEventsOnlyMode ? 'the bracket' : 'the main event'}. Should <strong>all {subEvents.length} sub-events</strong> adopt the same visibility (location filter + mailing lists + combination)?<br /><br />This usually makes sense so that everyone who should see the event can also reach the sub-events — access runs through the overall event’s visibility anyway. Any existing, differing sub-event visibilities will be <strong>overwritten</strong>.</>}
            </p>
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, flexWrap: 'wrap' }}>
              <button className="btn btn-secondary" onClick={() => closeVisCopy(false)}>
                {isDe ? 'Nein, eigene behalten' : 'No, keep their own'}
              </button>
              <button className="btn btn-primary" onClick={() => closeVisCopy(true)}>
                {isDe ? `Ja, auf alle ${subEvents.length} übernehmen` : `Yes, apply to all ${subEvents.length}`}
              </button>
            </div>
          </div>
        )}
      </Modal>

      {/* v11.57 / v11.63 / v13.4: Outlook-Update-Confirm-Modal über <Modal>-Wrapper.
          dismissable=false, da Schließen nur über Cancel-Button erlaubt. */}
      <Modal
        open={outlookConfirmOpen}
        onClose={cancelOutlookSave}
        maxWidth={620}
        dismissable={false}
        ariaLabel="Outlook-Update bestätigen"
      >
        {outlookConfirmOpen && (
          <div>
            <h2 id="outlook-confirm-title" style={{
              margin: '0 0 10px', fontSize: '1.15rem', fontWeight: 700,
              color: 'var(--dex-green-dark, #4a7c1f)',
            }}>
              {isDe ? 'Outlook-Termin der Teilnehmer aktualisieren?' : 'Update Outlook invite for attendees?'}
            </h2>
            <p style={{ margin: '0 0 14px', fontSize: '0.9rem', color: 'var(--dex-gray-700)', lineHeight: 1.55 }}>
              {isDe
                ? 'Du hast Felder geändert, die für die Teilnehmer-Outlook-Termine relevant sind. Wähle aus, welche Termine du jetzt neu rausschicken willst — der Rest wird gespeichert, aber Outlook bleibt unangetastet (du kannst das später jederzeit nachholen).'
                : 'You changed fields that are relevant to the attendees’ Outlook invites. Pick which invites you want to resend now — everything else is saved, but Outlook is left alone (you can resend later at any time).'}
            </p>
            <div style={{
              border: '1px solid var(--dex-gray-200)',
              borderRadius: 8,
              marginBottom: 14,
              background: 'var(--dex-gray-50, #f8f9fa)',
            }}>
              {outlookConfirmItems.map((it, idx) => {
                const isLast = idx === outlookConfirmItems.length - 1;
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
                const changedLabels = it.changedFields.map(f => isDe ? fieldLabelMap[f].de : fieldLabelMap[f].en).join(', ');
                const checked = !!outlookConfirmChecks[it.eventId];
                // v11.69: noOutlookYet-Items bekommen wieder eine Checkbox.
                // Default UNCHECKED. Beim Anhaken wird das Sub-Event in der
                // Eventverwaltung komplett neu angelegt (DEX_Events-Item
                // delete + create mit `existingSubsiteUrl`), damit der
                // Outlook-Termin entsteht. Die bestehende Teilnehmerliste
                // mit allen Anmeldungen bleibt unangetastet.
                if (it.noOutlookYet) {
                  return (
                    <label
                      key={it.eventId}
                      style={{
                        display: 'flex', alignItems: 'flex-start', gap: 12,
                        padding: '12px 14px',
                        borderBottom: isLast ? 'none' : '1px solid var(--dex-gray-200)',
                        cursor: 'pointer',
                        background: '#fffaf0',
                      }}
                    >
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={e => {
                          const next = e.target.checked;
                          setOutlookConfirmChecks(prev => ({ ...prev, [it.eventId]: next }));
                        }}
                        style={{ width: 18, height: 18, marginTop: 2, cursor: 'pointer', flexShrink: 0 }}
                      />
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: '0.92rem', fontWeight: 600, color: 'var(--dex-gray-800)', wordBreak: 'break-word' }}>
                          {isDe ? `Sub-Event: ${it.title}` : `Sub-event: ${it.title}`}
                        </div>
                        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 3 }}>
                          {isDe ? 'Geändert: ' : 'Changed: '}{changedLabels}
                        </div>
                        <div style={{ fontSize: '0.76rem', color: '#8a6d3b', marginTop: 6, lineHeight: 1.45, background: '#fcf8e3', border: '1px solid #faebcc', borderRadius: 4, padding: '6px 8px' }}>
                          {isDe
                            ? <>Für dieses Sub-Event gibt es noch keinen Outlook-Termin. Wenn du den Haken setzt, wird das Sub-Event in der Eventverwaltung neu angelegt, damit der Outlook-Termin entsteht. <strong>Die bestehende Teilnehmerliste mit allen Anmeldungen bleibt erhalten</strong> — nur die DEX_Events-Zeile bekommt eine neue ID.</>
                            : <>This sub-event does not have an Outlook event yet. If you tick the box, the sub-event is re-created in the event admin so the Outlook event can be generated. <strong>The existing participant list with all registrations stays intact</strong> — only the DEX_Events row gets a new ID.</>}
                        </div>
                      </div>
                    </label>
                  );
                }
                // v15.3: leere changedFields-Liste = Item kommt aus dem
                // persistierten OutlookDirty-Flag (frühere Session,
                // wurde damals nicht synchronisiert). Klartext-Hinweis
                // statt leerer „Geändert:"-Zeile.
                const isFromPersistedDirty = it.changedFields.length === 0;
                return (
                  <label
                    key={it.eventId}
                    style={{
                      display: 'flex', alignItems: 'flex-start', gap: 12,
                      padding: '12px 14px',
                      borderBottom: isLast ? 'none' : '1px solid var(--dex-gray-200)',
                      cursor: 'pointer',
                      background: isFromPersistedDirty ? '#fff8e8' : '#fff',
                    }}
                  >
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={e => {
                        const next = e.target.checked;
                        setOutlookConfirmChecks(prev => ({ ...prev, [it.eventId]: next }));
                      }}
                      style={{ width: 18, height: 18, marginTop: 2, cursor: 'pointer', flexShrink: 0 }}
                    />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: '0.92rem', fontWeight: 600, color: 'var(--dex-gray-800)', wordBreak: 'break-word' }}>
                        {it.kind === 'top'
                          ? (isDe ? `Hauptevent: ${it.title}` : `Main event: ${it.title}`)
                          : (isDe ? `Sub-Event: ${it.title}` : `Sub-event: ${it.title}`)}
                      </div>
                      {isFromPersistedDirty ? (
                        <div style={{ fontSize: '0.78rem', color: '#8a6d3b', marginTop: 4, lineHeight: 1.45 }}>
                          {isDe
                            ? <>⏳ <strong>Frühere Änderung nicht synchronisiert</strong> — beim letzten Speichern dieses Events wurden Outlook-relevante Felder geändert, der Outlook-Sync wurde aber damals übersprungen. Haken setzen, um die Teilnehmer jetzt nachträglich per Outlook-Update zu informieren.</>
                            : <>⏳ <strong>Earlier change not yet synced</strong> — Outlook-relevant fields were changed in a previous save of this event, but the Outlook sync was skipped at the time. Tick the box to send the catch-up Outlook update to attendees now.</>}
                        </div>
                      ) : (
                        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 3 }}>
                          {isDe ? 'Geändert: ' : 'Changed: '}{changedLabels}
                        </div>
                      )}
                    </div>
                  </label>
                );
              })}
            </div>
            <p style={{
              margin: '0 0 12px', fontSize: '0.8rem', color: 'var(--dex-gray-500)',
              lineHeight: 1.5,
            }}>
              {isDe
                ? 'Bei angehakten Events bekommen die Teilnehmer eine „Aktualisierter Termin"-Benachrichtigung von Outlook. Nicht angehakte Termine werden für später als „ausstehender Outlook-Sync" markiert.'
                : 'Ticked events trigger an “updated meeting” notification from Outlook for attendees. Unticked invites are flagged as “pending Outlook sync” for later.'}
            </p>
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 8 }}>
              <button
                className="btn btn-secondary"
                onClick={cancelOutlookSave}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                className="btn btn-primary"
                style={{ background: 'var(--dex-green, #86bc25)', borderColor: 'var(--dex-green, #86bc25)' }}
                onClick={() => confirmOutlookSave()}
              >
                {isDe ? 'Speichern' : 'Save'}
              </button>
            </div>
          </div>
        )}
      </Modal>

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
          maxWidth={480}
          padding={24}
          ariaLabel={isDe ? 'Ungespeicherte Änderungen' : 'Unsaved changes'}
        >
          <h3 style={{ margin: '0 0 12px', fontSize: '1.1rem', color: 'var(--dex-orange-dark, #b35a00)' }}>
            {isDe
              ? (isEditMode ? 'Ungespeicherte Änderungen' : 'Entwurf noch nicht gespeichert')
              : (isEditMode ? 'Unsaved changes' : 'Draft not saved yet')}
          </h3>
          <p style={{ margin: '0 0 16px', fontSize: '0.9rem', lineHeight: 1.5, color: 'var(--dex-gray-700)' }}>
            {isDe
              ? (isEditMode
                ? <>Du hast Änderungen am Event vorgenommen, die noch <strong>nicht gespeichert</strong> sind. Was möchtest du tun?</>
                : <>Dein Event ist noch <strong>nicht angelegt</strong>. Du kannst den Stand als Entwurf behalten — beim nächsten Öffnen der Event-Erstellung machst du genau hier weiter. Hochgeladene Bilder sind im Entwurf nicht enthalten.</>)
              : (isEditMode
                ? <>You have made changes to this event that are <strong>not saved yet</strong>. What do you want to do?</>
                : <>Your event is <strong>not created yet</strong>. You can keep this state as a draft — next time you open event creation you continue right here. Uploaded images are not part of the draft.</>)}
          </p>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
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
                style={{ fontSize: '0.9rem', width: '100%', justifyContent: 'center' }}
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
                style={{ fontSize: '0.9rem', width: '100%', justifyContent: 'center' }}
              >
                <Send size={14} /> {isDe ? 'Entwurf speichern' : 'Save draft'}
              </button>
            )}
            <button
              type="button"
              className="btn btn-secondary"
              onClick={() => { unsavedConfirmOpen.resolve(false); setUnsavedConfirmOpen(null); }}
              style={{ fontSize: '0.9rem', width: '100%', justifyContent: 'center' }}
            >
              {isDe
                ? (isEditMode ? 'Bearbeitung fortsetzen' : 'Eventerstellung fortsetzen')
                : (isEditMode ? 'Continue editing' : 'Continue creating')}
            </button>
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
              style={{ fontSize: '0.9rem', width: '100%', justifyContent: 'center' }}
            >
              {isDe
                ? (isEditMode ? 'Änderungen verwerfen' : 'Event verwerfen')
                : (isEditMode ? 'Discard changes' : 'Discard event')}
            </button>
          </div>
        </Modal>
      )}
    </>
  );
};
