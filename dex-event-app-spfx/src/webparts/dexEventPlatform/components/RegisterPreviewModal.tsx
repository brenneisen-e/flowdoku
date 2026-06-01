/**
 * Register-Page-Preview-Modal — v7.24 grosse Vorschau:
 *
 * Vorher: hand-gebaute Mockup-Cards (links Event-Info, rechts Form-Mock).
 * Jetzt: das ECHTE <RegistrationPage>-Component wird gerendert, mit einem
 * synthetischen DeloitteEvent aus dem Wizard-Form-State und den dazugehoerigen
 * Preview-Provider-Stubs (kein Live-API-Call). Wrapper hat pointer-events:none,
 * damit der Organizer schauen, aber nicht interagieren kann.
 *
 * Ziel: was der Organizer hier sieht, ist 1:1 das was der Teilnehmer im echten
 * Anmeldeformular bekommt — inklusive Sichtbarkeitsbedingungen, Multi-Select,
 * Roommate-Picker, helpText-Tooltips und Custom-Field-Reihenfolge.
 */

import * as React from 'react';
import { X } from './Icons';
import { PreviewContextStack } from './manual/previews/PreviewProviders';
import RegistrationPage from './RegistrationPage';
import Header from './Header';
import { DeloitteEvent } from '../types';

export interface RegisterPreviewData {
  title: string;
  description: string;
  location: string;
  locationAddress: { street?: string; houseNo?: string; zip?: string; city?: string };
  startDate: string;   // "YYYY-MM-DDTHH:mm" local
  endDate: string;
  imagePreview?: string;
  organizers: string[];
  organizerEmails: string[];
  maxParticipants: number;
  unlimitedParticipants: boolean;
  customFields: Array<{
    id: string;
    label: string;
    type: string;
    required: boolean;
    visible: boolean;
    options?: string[];
    helpText?: string;
    multi?: boolean;
    showIf?: { fieldId: string; values: string[] };
  }>;
  isFictive?: boolean;
  /** v14.10: Sub-Event-Drafts für die Vorschau — damit der Organizer im
   *  Anmeldeformular die Sub-Event-Liste sieht (vorher fehlte sie komplett). */
  subEvents?: Array<{
    id: string;
    title: string;
    location?: string;
    startDate: string;
    endDate: string;
    maxParticipants?: number;
    description?: string;
    customFields?: Array<{
      id: string;
      label: string;
      type: string;
      required: boolean;
      visible?: boolean;
      options?: string[];
      helpText?: string;
      multi?: boolean;
      showIf?: { fieldId: string; values: string[] };
    }>;
  }>;
  /** v14.10: subEventsOnlyMode + childEventTerm zur korrekten Darstellung
   *  in der Vorschau (Pflicht-Hinweis, Anmelde-Checkbox-Sichtbarkeit). */
  subEventsOnlyMode?: boolean;
  requireSubEventSelection?: boolean;
  childEventTermSingular?: string;
  childEventTermPlural?: string;
  /** v17.22: Bilingual-Toggle an die Vorschau weiterreichen, damit
   *  RegistrationPage die EN-Varianten der Felder rendert (useEnVariants
   *  greift nur, wenn event.bilingualFields gesetzt ist). */
  bilingualFields?: boolean;
}

export interface RegisterPreviewModalProps {
  open: boolean;
  onClose: () => void;
  data: RegisterPreviewData;
}

// Synthetischen DeloitteEvent aus dem Wizard-Form-State bauen, sodass
// RegistrationPage ihn ueber den EventContext finden kann.
const SYNTH_PARENT_ID = '__preview_synth_event__';

function buildSynthEvent(data: RegisterPreviewData): DeloitteEvent {
  return {
    id: SYNTH_PARENT_ID,
    eventNumber: 0,
    title: data.title || '(kein Titel)',
    type: 'Other',
    status: 'Active',
    organizers: data.organizers || [],
    organizerEmails: data.organizerEmails || [],
    qrScannerNames: [],
    qrScannerEmails: [],
    location: data.location || '',
    locationAddress: data.locationAddress
      ? {
          street: data.locationAddress.street || '',
          houseNo: data.locationAddress.houseNo || '',
          zip: data.locationAddress.zip || '',
          city: data.locationAddress.city || '',
        }
      : undefined,
    locationAudience: [],
    audienceFilter: [],
    filterMode: 'OR',
    startDate: data.startDate ? new Date(data.startDate).toISOString() : '',
    endDate: data.endDate ? new Date(data.endDate).toISOString() : '',
    registrationDeadline: '',
    lastDeregisterDate: '',
    description: data.description || '',
    maxParticipants: data.unlimitedParticipants ? 0 : (data.maxParticipants || 0),
    currentParticipants: 0,
    waitlistCount: 0,
    imageUrl: data.imagePreview || '',
    outlookBody: '',
    emailLanguage: 'DE',
    bilingualFields: !!data.bilingualFields,
    isFictive: !!data.isFictive,
    subEventsOnlyMode: !!data.subEventsOnlyMode,
    requireSubEventSelection: !!data.requireSubEventSelection,
    childEventTermSingular: data.childEventTermSingular,
    childEventTermPlural: data.childEventTermPlural,
    agenda: [],
    transferTimes: [],
    documents: [],
    quiz: [],
    eventSpecificFields: (data.customFields || [])
      .filter(f => f.visible !== false && f.label && f.label.trim().length > 0)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .map((f: any) => ({
        id: f.id,
        label: f.label,
        type: (f.type as 'text') || 'text',
        required: !!f.required,
        options: f.options,
        helpText: f.helpText || '',
        helpTextStyle: f.helpTextStyle === 'inline' ? 'inline' : 'tooltip',
        ...(f.multi ? { multi: true } : {}),
        ...(f.showIf ? { showIf: f.showIf } : {}),
        // v17.22: EN-Varianten + confirmLabel an den Synth-Event durchreichen,
        // damit die Vorschau im EN-Modus die englischen Texte rendert.
        ...(f.confirmLabel ? { confirmLabel: f.confirmLabel } : {}),
        ...(f.labelEn ? { labelEn: f.labelEn } : {}),
        ...(f.helpTextEn ? { helpTextEn: f.helpTextEn } : {}),
        ...(f.confirmLabelEn ? { confirmLabelEn: f.confirmLabelEn } : {}),
        ...(f.optionsEn ? { optionsEn: f.optionsEn } : {}),
      })),
  };
}

// v14.10: Sub-Event-Drafts als synthetische DeloitteEvent-Children bauen,
// damit die RegistrationPage sie über childEventsOf(parentId) findet.
function buildSynthChildEvents(data: RegisterPreviewData): DeloitteEvent[] {
  return (data.subEvents || [])
    .filter(s => s.title && s.title.trim().length > 0)
    .map((s, idx) => ({
      id: `${SYNTH_PARENT_ID}_child_${idx}`,
      parentEventId: SYNTH_PARENT_ID,
      eventNumber: idx + 1,
      title: s.title,
      type: 'Other',
      status: 'Active',
      organizers: data.organizers || [],
      organizerEmails: data.organizerEmails || [],
      qrScannerNames: [],
      qrScannerEmails: [],
      location: s.location || '',
      locationAudience: [],
      audienceFilter: [],
      filterMode: 'OR',
      startDate: s.startDate ? new Date(s.startDate).toISOString() : '',
      endDate: s.endDate ? new Date(s.endDate).toISOString() : '',
      registrationDeadline: '',
      lastDeregisterDate: '',
      description: s.description || '',
      maxParticipants: s.maxParticipants || 0,
      currentParticipants: 0,
      waitlistCount: 0,
      imageUrl: '',
      outlookBody: '',
      emailLanguage: 'DE',
      bilingualFields: !!data.bilingualFields,
      isFictive: !!data.isFictive,
      agenda: [],
      transferTimes: [],
      documents: [],
      quiz: [],
      eventSpecificFields: (s.customFields || [])
        .filter(f => f.visible !== false && f.label && f.label.trim().length > 0)
        .map(f => ({
          id: f.id,
          label: f.label,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          type: (f.type as any) || 'text',
          required: !!f.required,
          options: f.options,
          helpText: f.helpText || '',
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          helpTextStyle: ((f as any).helpTextStyle === 'inline' ? 'inline' : 'tooltip') as 'inline' | 'tooltip',
          ...(f.multi ? { multi: true } : {}),
          ...(f.showIf ? { showIf: f.showIf } : {}),
        })),
    }));
}

export const RegisterPreviewModal: React.FC<RegisterPreviewModalProps> = ({ open, onClose, data }) => {
  if (!open) return null;
  const synthEvent = React.useMemo(() => buildSynthEvent(data), [data]);
  const synthChildren = React.useMemo(() => buildSynthChildEvents(data), [data]);

  return (
    <div
      role="dialog"
      aria-modal="true"
      onClick={onClose}
      style={{
        position: 'fixed', inset: 0, zIndex: 1300,
        background: 'rgba(0,0,0,0.55)',
        display: 'flex', alignItems: 'stretch', justifyContent: 'center',
        padding: '24px 16px', overflowY: 'auto',
      }}
    >
      <div
        onClick={e => e.stopPropagation()}
        style={{
          width: '100%', maxWidth: 1280,
          background: '#fff', borderRadius: 'var(--dex-radius, 12px)',
          boxShadow: '0 20px 60px rgba(0,0,0,0.3)',
          display: 'flex', flexDirection: 'column', overflow: 'hidden',
          maxHeight: 'calc(100vh - 48px)',
        }}
      >
        {/* Header */}
        <div style={{
          display: 'flex', alignItems: 'center', justifyContent: 'space-between',
          padding: '14px 20px', borderBottom: '1px solid var(--dex-gray-200)',
          background: '#fff', flexShrink: 0,
        }}>
          <h3 style={{ margin: 0, fontSize: '1.05rem' }}>
            Registrierungsseite — Vorschau
            {data.isFictive && (
              <span style={{ marginLeft: 10, padding: '2px 10px', borderRadius: 999, background: 'var(--dex-orange, #ed8b00)', color: '#fff', fontSize: '0.7rem', fontWeight: 700, letterSpacing: 0.5 }}>
                ENTWURF
              </span>
            )}
            <span style={{ marginLeft: 10, fontSize: '0.7rem', fontWeight: 500, color: 'var(--dex-gray-500)' }}>
              · Vorschau · echte Anmeldung deaktiviert
            </span>
          </h3>
          <button
            type="button"
            onClick={onClose}
            style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 4 }}
            aria-label="Schließen"
          >
            <X size={20} />
          </button>
        </div>

        {/* Laptop-Frame (Desktop-Look): subtle dark bezel + subtle keyboard-base
            damit auf einen Blick klar ist dass das die Desktop-Ansicht ist. */}
        <div style={{
          flex: 1,
          padding: '24px 24px 36px',
          overflow: 'auto',
          background: 'var(--dex-gray-100, #f4f4f5)',
          display: 'flex', justifyContent: 'center',
        }}>
          <div style={{ width: '100%', maxWidth: 1200, position: 'relative', paddingBottom: 14 }}>
            <div style={{
              padding: '14px 14px 12px', borderRadius: '12px 12px 6px 6px',
              background: '#1a1a1a',
              boxShadow: '0 8px 28px rgba(0,0,0,0.2)',
              position: 'relative',
            }}>
              <div style={{
                position: 'absolute', top: 5, left: '50%', transform: 'translateX(-50%)',
                width: 6, height: 6, borderRadius: '50%', background: '#333',
              }} />
              <div
                className="dex-preview-scope"
                // v14.10: Klicks sind jetzt erlaubt (Organizer kann durchklicken,
                // Sub-Events anhaken, Felder ausfüllen) — der EventContext-Stub
                // in PreviewContextStack hat `registerForEvent` als asyncNoop,
                // ein echter Submit passiert also nie. Nur Form-Submit-Events
                // abfangen, damit kein Page-Reload getriggert wird.
                onSubmitCapture={e => { e.preventDefault(); e.stopPropagation(); }}
                style={{
                  background: '#fff', borderRadius: 4,
                  pointerEvents: 'auto', userSelect: 'text',
                  minHeight: 560, maxHeight: '70vh', overflow: 'auto',
                }}
              >
                <PreviewContextStack
                  role="User"
                  page="register"
                  selectedEventId={synthEvent.id}
                  extraEvents={[synthEvent, ...synthChildren]}
                >
                  <Header />
                  <RegistrationPage />
                </PreviewContextStack>
              </div>
            </div>
            {/* Laptop-Stand */}
            <div style={{
              position: 'absolute', left: -12, right: -12, bottom: 0, height: 12,
              background: 'linear-gradient(180deg, #2a2a2a, #111)',
              borderRadius: '0 0 12px 12px',
            }} />
            <div style={{
              position: 'absolute', left: '40%', right: '40%', bottom: 8, height: 4,
              background: '#444', borderRadius: '0 0 6px 6px',
            }} />
          </div>
        </div>

        <div style={{
          display: 'flex', justifyContent: 'flex-end', gap: 8,
          padding: '12px 20px', borderTop: '1px solid var(--dex-gray-200)',
          background: '#fff', flexShrink: 0,
        }}>
          <button type="button" className="btn btn-primary" onClick={onClose}>Schließen</button>
        </div>
      </div>
    </div>
  );
};
