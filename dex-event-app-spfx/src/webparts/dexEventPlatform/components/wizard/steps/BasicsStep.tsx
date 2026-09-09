/* BasicsStep — aus EventCreationPage.tsx ausgelagert (Zeilen 9045-10108 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; einzige
 * Aenderung ist die Anzeige-Bedingung: aus `currentStep === 0` wurde das Prop `visible`.
 * `visible` schaltet display:none statt unmount — Eingaben ueberleben den
 * Schrittwechsel genauso wie vorher. */
import * as React from 'react';
import WizardHint from '../../WizardHint';
import { StepBadge } from '../../wizard/StepBadge';
import { InfoTooltip } from '../../InfoTooltip';
import DatePicker from 'react-datepicker';
import { Check, Pencil, Plus, X } from '../../Icons';
import { Icon } from '@fluentui/react/lib/Icon';
// v31.2: gemeinsame UI-Klassen (Toggle-Zeilen, Kacheln, Aufklapper) — Hover
// kommt aus dem Stylesheet, nicht aus Inline-Styles.
import { cx } from '../../dexUi';
import ImageCropModal from '../../ImageCropModal';
import { compressImage } from '../../../utils/imageCompress';
import { ImgView, SubEventDraft } from '../../wizard/wizardTypes';
export interface BasicsStepProps {
  visible: boolean;
  activeFrom: string;
  activeScopeIdx: number;
  applyDraftPayload: (d: Record<string, unknown>) => void;
  applyEventTemplate: (ev: import("../../../types/index").DeloitteEvent) => Promise<void>;
  childEventsOf: (parentEventId: string) => import("../../../types/index").DeloitteEvent[];
  childTermSingular: string;
  currentUser: import("../../../types/index").User;
  dayKeyOfDate: (d: Date) => string;
  description: string;
  DRAFT_KEY: string;
  draftSavedAt: number;
  editEvent: import("../../../types/index").DeloitteEvent;
  emailLogoPreview: string;
  errorBorderStyle: (fieldName: string) => React.CSSProperties;
  events: import("../../../types/index").DeloitteEvent[];
  fieldHasError: (fieldName: string) => boolean;
  fileToBase64: (file: File) => Promise<string>;
  imageBanner: boolean;
  imageDisplay: { card?: ImgView; hero?: ImgView; };
  imageDisplayOpen: boolean;
  imageEditOpen: boolean;
  imageFile: File;
  imageOrigFile: File;
  imagePreview: string;
  imageUploadError: string;
  isDe: boolean;
  isEditMode: boolean;
  isFictive: boolean;
  location: string;
  logoCropTarget: "outlook" | "email";
  noDescription: boolean;
  outlookLogoPreview: string;
  patchScopeSub: (patch: Partial<SubEventDraft>) => void;
  pendingDraft: { savedAt: number; data: Record<string, unknown>; };
  previewBeforeActive: boolean;
  renderStepIntro: (_bulletsDe: string[], _bulletsEn: string[]) => React.ReactElement | null;
  scAllDay: boolean;
  scDescription: string;
  scEnd: Date;
  scImagePreview: string;
  scopeSub: SubEventDraft;
  scShowAsFree: boolean;
  scStart: Date;
  scTitle: string;
  setActiveFrom: React.Dispatch<React.SetStateAction<string>>;
  setDescription: React.Dispatch<React.SetStateAction<string>>;
  setEmailLogoFromPhoto: React.Dispatch<React.SetStateAction<boolean>>;
  setEmailLogoPreview: React.Dispatch<React.SetStateAction<string>>;
  setEventImageUrl: React.Dispatch<React.SetStateAction<string>>;
  setHtmlEditorMode: React.Dispatch<React.SetStateAction<"outlook" | "email" | "description">>;
  setHtmlEditorOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setImageBanner: React.Dispatch<React.SetStateAction<boolean>>;
  setImageDisplay: React.Dispatch<React.SetStateAction<{ card?: ImgView; hero?: ImgView; }>>;
  setImageDisplayOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setImageEditOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setImageFile: React.Dispatch<React.SetStateAction<File>>;
  setImageOrigAspect: React.Dispatch<React.SetStateAction<number>>;
  setImageOrigFile: React.Dispatch<React.SetStateAction<File>>;
  setImagePreview: React.Dispatch<React.SetStateAction<string>>;
  setImageUploadError: React.Dispatch<React.SetStateAction<string>>;
  setIsFictive: React.Dispatch<React.SetStateAction<boolean>>;
  setLogoCropTarget: React.Dispatch<React.SetStateAction<"outlook" | "email">>;
  setNoDescription: React.Dispatch<React.SetStateAction<boolean>>;
  setOutlookLogoFromPhoto: React.Dispatch<React.SetStateAction<boolean>>;
  setOutlookLogoPreview: React.Dispatch<React.SetStateAction<string>>;
  setPendingDraft: React.Dispatch<React.SetStateAction<{ savedAt: number; data: Record<string, unknown>; }>>;
  setPreviewBeforeActive: React.Dispatch<React.SetStateAction<boolean>>;
  setScAllDay: (v: boolean) => void;
  setScEnd: (d: Date | null) => void;
  setScShowAsFree: (v: boolean) => void;
  setScStart: (d: Date | null) => void;
  setScTitle: (v: string) => void;
  setShowDemoVariantModal: React.Dispatch<React.SetStateAction<boolean>>;
  setShowTemplatePicker: React.Dispatch<React.SetStateAction<boolean>>;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setSubImageCropIdx: React.Dispatch<React.SetStateAction<number>>;
  showTemplatePicker: boolean;
  shrinkLogoB64: (b64: string) => Promise<string>;
  startDate: string;
  subEvents: SubEventDraft[];
  subEventsOnlyMode: boolean;
  t: (key: string) => string;
  templateLoadingId: string;
  title: string;
  wizardImgAspect: number;
  zebraS3Bg: () => string;
}
export const BasicsStep: React.FC<BasicsStepProps> = (p) => {
  const { visible } = p;
  const { activeFrom, activeScopeIdx, applyDraftPayload, applyEventTemplate, childEventsOf, childTermSingular, currentUser, dayKeyOfDate, description, DRAFT_KEY, draftSavedAt, editEvent, emailLogoPreview, errorBorderStyle, events, fieldHasError, fileToBase64, imageBanner, imageDisplay, imageDisplayOpen, imageEditOpen, imageFile, imageOrigFile, imagePreview, imageUploadError, isDe, isEditMode, isFictive, location, logoCropTarget, noDescription, outlookLogoPreview, patchScopeSub, pendingDraft, previewBeforeActive, renderStepIntro, scAllDay, scDescription, scEnd, scImagePreview, scopeSub, scShowAsFree, scStart, scTitle, setActiveFrom, setDescription, setEmailLogoFromPhoto, setEmailLogoPreview, setEventImageUrl, setHtmlEditorMode, setHtmlEditorOpen, setImageBanner, setImageDisplay, setImageDisplayOpen, setImageEditOpen, setImageFile, setImageOrigAspect, setImageOrigFile, setImagePreview, setImageUploadError, setIsFictive, setLogoCropTarget, setNoDescription, setOutlookLogoFromPhoto, setOutlookLogoPreview, setPendingDraft, setPreviewBeforeActive, setScAllDay, setScEnd, setScShowAsFree, setScStart, setScTitle, setShowDemoVariantModal, setShowTemplatePicker, setSubEvents, setSubImageCropIdx, showTemplatePicker, shrinkLogoB64, startDate, subEvents, subEventsOnlyMode, t, templateLoadingId, title, wizardImgAspect } = p;
  // v31.2 (Leitfaden 2a′): Eine Kachel mit Hauptaktion ist selbst klickbar —
  // Enter/Leertaste lösen dieselbe Aktion aus wie der Klick. Nebenknöpfe in
  // der Kachel stoppen die Weitergabe (Klick UND Taste), damit „Verwerfen"
  // nicht zugleich „Fortsetzen" ist.
  const rowKeyHandler = (fn: () => void) => (e: React.KeyboardEvent<HTMLElement>): void => {
    if (e.target !== e.currentTarget) return;
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fn(); }
  };
  const stopBubble = (e: React.SyntheticEvent): void => { e.stopPropagation(); };
  // Ein Handler für Knopf UND Kachel — sonst zeigen beide irgendwann
  // verschiedene Dinge.
  const openDescriptionEditor = (): void => { setHtmlEditorMode('description'); setHtmlEditorOpen(true); };
  return (
              <div style={{ display: visible ? 'block' : 'none' }}>
              {/* v23.6: Demo-Button sitzt jetzt IM grünen Schritt-1-Header
                  (oben rechts), nicht mehr in einer eigenen Zeile darüber. */}
              {/* v31.2: Kopf nach Leitfaden (Eyebrow + Titel). Der Demo-Chip
                  steht nach 2a′ direkt HINTER dem Titel statt allein am rechten
                  Rand — ein Knopf in einer sonst leeren Zeilenhälfte war der
                  Fehler aus den ersten Screenshots. */}
              <h2 className="dex-step-head-title" style={{ flexDirection: 'row', alignItems: 'flex-end', justifyContent: 'flex-start', gap: 12, flexWrap: 'wrap' }}>
                <span style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                  <span className="dex-step-eyebrow">{isDe ? 'Schritt 1 von 9' : 'Step 1 of 9'}</span>
                  <span>{isDe ? 'Grundlagen' : 'Basics'}</span>
                </span>
                {!isEditMode && (
                  <button
                    type="button"
                    className="dex-ui-chip"
                    data-tour="wizard-demo"
                    onClick={() => setShowDemoVariantModal(true)}
                    title={isDe ? 'Demo-Vorlage auswählen' : 'Choose demo template'}
                    style={{ flexShrink: 0, marginBottom: 4 }}
                  >
                    {isDe ? 'Demo-Vorlage' : 'Demo template'}
                  </button>
                )}
              </h2>
              <p className="dex-step-head-lead">
                {isDe
                  ? <>Titel, Zeitraum, Beschreibung und Bild — das, was Teilnehmer zuerst sehen. Ganz unten legst du fest, <strong>wann das Event sichtbar wird</strong>.</>
                  : <>Title, dates, description and image — what attendees see first. At the bottom you decide <strong>when the event becomes visible</strong>.</>}
              </p>

              {/* v28.89: Alles zwischen hier und dem Titel-Feld gilt für das
                  GESAMTE Event — ob es Sub-Events gibt, wie sie heißen, wie
                  angemeldet wird, Entwurf/Aktivierung. Auf einem Sub-Event-
                  Reiter wäre das falsch am Platz (man würde die Grundsatzfrage
                  „unter" einem einzelnen Termin beantworten), deshalb blenden
                  wir es dort aus und sagen, wo es steht. */}
              {activeScopeIdx > 0 && (
                <div className="dex-ui-callout dex-ui-callout--info" style={{ marginBottom: 16 }}>
                  <span className="dex-ui-callout-icon" aria-hidden="true"><Icon iconName="Info" style={{ fontSize: 14 }} /></span>
                  <span>
                    {isDe
                      ? <>Du bearbeitest die Grundlagen eines <strong>{childTermSingular || 'Sub-Events'}</strong>: Titel, Zeiten, Beschreibung und Bild unten gehören zu ihm. Alles, was das gesamte Event betrifft — ob es Sub-Events gibt, wie sie heißen, wie angemeldet wird sowie Entwurf und Aktivierung — steht auf dem Reiter <strong>{subEventsOnlyMode ? 'Klammer' : 'Haupt-Event'}</strong> oben.</>
                      : <>You are editing the basics of a <strong>{childTermSingular || 'sub-event'}</strong>: title, times, description and image below belong to it. Everything about the event as a whole — whether it has sub-events, how they are named, how people register, plus draft and activation — lives on the <strong>{subEventsOnlyMode ? 'bracket' : 'main event'}</strong> tab above.</>}
                  </span>
                </div>
              )}
              {activeScopeIdx === 0 && (<>




              {/* v24.9 (E): „Eigenes Event als Vorlage" — prominenter Fächer aus
                  Bildern bisheriger Events. Nur im NEU-Modus, nur wenn der
                  Organizer schon eigene Events hat. */}
              {!isEditMode && (() => {
                const meLc = (currentUser?.email || '').toLowerCase();
                const tmpl = (events || []).filter(e => {
                  if (e.parentEventId || e.isDemoShowcase) return false;
                  return (e.organizerEmails || []).some(x => (x || '').toLowerCase() === meLc)
                    || (e.coOrganizerEmails || []).some(x => (x || '').toLowerCase() === meLc);
                }).sort((a, b) => {
                  const ai = a.imageUrl ? 1 : 0, bi = b.imageUrl ? 1 : 0;
                  if (ai !== bi) return bi - ai;
                  const at = a.startDate ? new Date(a.startDate).getTime() : 0;
                  const bt = b.startDate ? new Date(b.startDate).getTime() : 0;
                  return bt - at;
                });
                if (tmpl.length === 0) return null;
                const fan = tmpl.filter(e => e.imageUrl).slice(0, 5);
                const fanItems = fan.length > 0 ? fan : tmpl.slice(0, 5);
                return (
                  <div className={cx('dex-ui-card', !showTemplatePicker && 'dex-ui-card--hover')} style={{ margin: '0 0 16px', padding: 0, overflow: 'hidden' }}>
                    {!showTemplatePicker ? (
                      <button
                        type="button"
                        onClick={() => setShowTemplatePicker(true)}
                        style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 20, padding: '16px 20px', background: 'transparent', border: 'none', cursor: 'pointer', textAlign: 'left', fontFamily: 'inherit' }}
                      >
                        {/* Fächer aus Event-Bildern */}
                        <div style={{ position: 'relative', width: 132, height: 92, flexShrink: 0 }}>
                          {fanItems.map((e, i) => {
                            const n = fanItems.length;
                            const spread = 16; // Grad pro Karte
                            const rot = (i - (n - 1) / 2) * spread;
                            const tx = (i - (n - 1) / 2) * 22;
                            return (
                              <div key={e.id} style={{
                                position: 'absolute', left: '50%', top: 6, width: 64, height: 80, marginLeft: -32,
                                borderRadius: 10, border: '3px solid #fff', boxShadow: '0 4px 10px rgba(0,0,0,0.18)',
                                transform: `translateX(${tx}px) rotate(${rot}deg)`, transformOrigin: 'bottom center',
                                background: e.imageUrl ? `url(${e.imageUrl}) center/cover no-repeat` : 'linear-gradient(135deg, var(--dex-green, #86bc25), var(--dex-blue, #0076a8))',
                                zIndex: i,
                              }} />
                            );
                          })}
                        </div>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontWeight: 700, fontSize: '0.95rem', color: 'var(--dex-gray-800)', marginBottom: 2 }}>
                            {isDe ? 'Eigenes Event als Vorlage nutzen?' : 'Use one of your events as a template?'}
                          </div>
                          <div className="dex-ui-muted" style={{ lineHeight: 1.5 }}>
                            {isDe
                              ? <>Übernimm Einstellungen und Bild aus einem deiner <strong>{tmpl.length}</strong> bisherigen Events — Datum und Anmeldungen legst du danach neu fest.</>
                              : <>Reuse settings and image from one of your <strong>{tmpl.length}</strong> past events — date and registrations start fresh.</>}
                          </div>
                          {/* v31.2 (2a′): Aktion linksbündig unter dem Text, nicht
                              rechts außen — die ganze Kachel ist ohnehin der Knopf. */}
                          <span className="dex-ui-textbtn" style={{ display: 'inline-flex', marginTop: 6 }}>{isDe ? 'Auswählen ▸' : 'Choose ▸'}</span>
                        </div>
                      </button>
                    ) : (
                      <div style={{ padding: '16px 20px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, marginBottom: 12 }}>
                          <span style={{ fontWeight: 700, fontSize: '0.95rem', color: 'var(--dex-gray-800)' }}>
                            {isDe ? 'Welches Event soll als Vorlage dienen?' : 'Which event should serve as the template?'}
                          </span>
                          <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" onClick={() => setShowTemplatePicker(false)}>
                            {isDe ? 'Abbrechen' : 'Cancel'}
                          </button>
                        </div>
                        <div style={{ display: 'flex', gap: 12, overflowX: 'auto', paddingBottom: 6 }}>
                          {tmpl.map(e => (
                            <button
                              key={e.id}
                              type="button"
                              className="dex-ui-card dex-ui-card--hover"
                              disabled={!!templateLoadingId}
                              onClick={() => { void applyEventTemplate(e); }}
                              style={{
                                flex: '0 0 auto', width: 150, textAlign: 'left', cursor: templateLoadingId ? 'wait' : 'pointer',
                                borderRadius: 12, padding: 0, overflow: 'hidden', fontFamily: 'inherit',
                              }}
                            >
                              <div style={{ width: '100%', height: 90, background: e.imageUrl ? `url(${e.imageUrl}) center/cover no-repeat` : 'linear-gradient(135deg, var(--dex-green, #86bc25), var(--dex-blue, #0076a8))' }} />
                              <div style={{ padding: '8px 10px' }}>
                                <div style={{ fontWeight: 700, fontSize: '0.82rem', color: 'var(--dex-gray-800)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                  {templateLoadingId === e.id ? (isDe ? 'Wird geladen…' : 'Loading…') : e.title}
                                </div>
                                {e.startDate && (
                                  <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                                    {new Date(e.startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })}
                                  </div>
                                )}
                              </div>
                            </button>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                );
              })()}

              {/* v30.4: Aktueller Entwurf als eigene Kachel unter der
                  Vorlagen-Kachel — statt des Modal-Dialogs beim Öffnen.
                  Verschwindet, sobald der User fortsetzt, löscht oder durch
                  eigenes Tippen einen neuen Autosave erzeugt (draftSavedAt). */}
              {!isEditMode && pendingDraft && draftSavedAt === null && (() => {
                const when = new Date(pendingDraft.savedAt).toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
                const dTitle = (typeof pendingDraft.data.title === 'string' && (pendingDraft.data.title as string).trim())
                  ? (pendingDraft.data.title as string).trim()
                  : (isDe ? '(ohne Titel)' : '(untitled)');
                const continueDraft = (): void => {
                  try { applyDraftPayload(pendingDraft.data); } catch (err) { console.warn('[DEX] Entwurf-Wiederherstellung fehlgeschlagen:', err); }
                  setPendingDraft(null);
                };
                // v31.2 (2a′): Die ganze Kachel öffnet den Entwurf (Hover,
                // Tastatur); die Knöpfe stehen linksbündig unter dem Text.
                // „Verwerfen" stoppt die Weitergabe, sonst würde es zugleich
                // fortsetzen.
                return (
                  <div
                    className="dex-ui-card dex-ui-card--hover dex-ui-fade-in"
                    role="button"
                    tabIndex={0}
                    onClick={continueDraft}
                    onKeyDown={rowKeyHandler(continueDraft)}
                    style={{ margin: '0 0 16px', borderColor: 'var(--dex-orange, #ed8b00)', background: 'rgba(237,139,0,0.05)', display: 'flex', alignItems: 'flex-start', gap: 16, flexWrap: 'wrap', cursor: 'pointer' }}
                  >
                    <span className="dex-ui-choice-icon" style={{ background: 'rgba(237,139,0,0.14)', color: 'var(--dex-orange-dark, #b35a00)' }} aria-hidden="true"><Pencil size={18} /></span>
                    <div style={{ flex: 1, minWidth: 220 }}>
                      <div style={{ fontWeight: 700, fontSize: '0.95rem', color: 'var(--dex-orange-dark, #b35a00)', marginBottom: 2 }}>
                        {isDe ? 'Du hast einen unfertigen Entwurf' : 'You have an unfinished draft'}
                      </div>
                      <div className="dex-ui-muted" style={{ lineHeight: 1.5 }}>
                        {isDe
                          ? <><strong>&bdquo;{dTitle}&ldquo;</strong> — zwischengespeichert am {when}. Hochgeladene Bilder sind im Entwurf nicht enthalten und müssten neu gewählt werden.</>
                          : <><strong>&bdquo;{dTitle}&ldquo;</strong> — auto-saved on {when}. Uploaded images are not part of the draft and would need to be re-selected.</>}
                      </div>
                      <div className="dex-ui-inline" style={{ marginTop: 10, gap: 12 }}>
                        <button
                          type="button"
                          className="btn btn-primary dex-ui-btn-sm"
                          onClick={e => { e.stopPropagation(); continueDraft(); }}
                          onKeyDown={stopBubble}
                        >
                          {isDe ? 'Entwurf fortsetzen' : 'Continue draft'}
                        </button>
                        <button
                          type="button"
                          className="dex-ui-textbtn dex-ui-textbtn--danger"
                          onClick={e => {
                            e.stopPropagation();
                            try { localStorage.removeItem(DRAFT_KEY); } catch { /* */ }
                            setPendingDraft(null);
                          }}
                          onKeyDown={stopBubble}
                        >
                          {isDe ? 'Entwurf verwerfen' : 'Discard draft'}
                        </button>
                      </div>
                    </div>
                  </div>
                );
              })()}

              {/* v11.57 / v11.63: Hinweisbox bei ausstehendem Outlook-Sync.
                  Sichtbar bei editEvent, wenn OutlookDirty=true auf dem
                  Hauptevent ODER auf mindestens einem Sub-Event gesetzt ist
                  (und Outlook für das jeweilige Event nicht deaktiviert
                  wurde). Auf neuen Events nie. */}
              {(() => {
                if (!editEvent) return null;
                // v18.51: im „Nur für Sub-Events"-Modus kein Hauptevent-Dirty-Hinweis.
                const topDirty = editEvent.outlookDirty === true && editEvent.disableOutlook !== true && !editEvent.subEventsOnlyMode;
                const dirtySubs = childEventsOf(editEvent.id).filter(k => k.outlookDirty === true && k.disableOutlook !== true);
                if (!topDirty && dirtySubs.length === 0) return null;
                const subCount = dirtySubs.length;
                let bodyDe: string;
                let bodyEn: string;
                if (topDirty && subCount > 0) {
                  bodyDe = `Outlook-Synchronisation steht aus: für das Hauptevent UND ${subCount} Sub-Event${subCount === 1 ? '' : 's'}. Beim nächsten Speichern kannst du im Dialog pro Termin entscheiden, ob die Teilnehmer eine „Aktualisierter Termin"-Benachrichtigung bekommen.`;
                  bodyEn = `Outlook sync is pending: for the main event AND ${subCount} sub-event${subCount === 1 ? '' : 's'}. On the next save you can decide per invite whether the attendees should receive an “updated meeting” notification.`;
                } else if (topDirty) {
                  bodyDe = 'Outlook-Synchronisation steht aus: für das Hauptevent. Beim nächsten Speichern kannst du im Dialog pro Termin entscheiden, ob die Teilnehmer eine „Aktualisierter Termin"-Benachrichtigung bekommen.';
                  bodyEn = 'Outlook sync is pending: for the main event. On the next save you can decide per invite whether the attendees should receive an “updated meeting” notification.';
                } else {
                  bodyDe = `Outlook-Synchronisation steht aus: für ${subCount} Sub-Event${subCount === 1 ? '' : 's'}. Beim nächsten Speichern kannst du im Dialog pro Termin entscheiden, ob die Teilnehmer eine „Aktualisierter Termin"-Benachrichtigung bekommen.`;
                  bodyEn = `Outlook sync is pending: for ${subCount} sub-event${subCount === 1 ? '' : 's'}. On the next save you can decide per invite whether the attendees should receive an “updated meeting” notification.`;
                }
                return (
                  <WizardHint
                    isDe={isDe}
                    title={isDe ? 'Outlook-Synchronisation steht aus' : 'Outlook sync is pending'}
                    style={{ marginBottom: 16 }}
                  >
                    {isDe ? bodyDe : bodyEn}
                  </WizardHint>
                );
              })()}

              {renderStepIntro(
                [
                  '5. Als Entwurf speichern (Abschnitt „Veröffentlichung" ganz unten) — Event nur für Admins, Organizer und Test-Team sichtbar; optional Aktiv-Ab-Datum für automatisches Go-Live',
                  '1. Event-Titel',
                  '2. Datum (Start &amp; Ende) — füllt die Anmelde- und Storno-Deadlines automatisch vor',
                  '3. Beschreibung (optional, HTML-Editor)',
                  '4. Event-Bild hochladen — oben auf der Detailseite und in den Mails verwendet',
                  '6. Organizer auswählen — bekommen alle Organizer-Mails',
                  '7. Test-Team — sieht das Event schon im Entwurfsmodus',
                  '8. Check-In Team — darf nur das QR-/Check-In-Tool nutzen',
                ],
                [
                  '5. Save as draft (section “Publishing” at the bottom) — visible only to admins, organizers, and the test team; optional active-from date for automatic go-live',
                  '1. Event title',
                  '2. Date (start &amp; end) — pre-fills the registration and cancellation deadlines',
                  '3. Description (optional, HTML editor)',
                  '4. Upload an event image — shown at the top of the detail page and in emails',
                  '6. Pick organizers — they receive all organizer emails',
                  '7. Test team — can see the event already in draft mode',
                  '8. Check-in team — may only use the QR / check-in tool',
                ]
              )}

              {/* v31.2: Der Entwurfs-/Aktivierungs-Block (Badge 5) steht jetzt
                  als Abschnitt „Veröffentlichung" am Ende des Schritts. */}
              </>)}{/* v28.89: Ende der event-weiten Angaben */}

              {/* v31.2: Jeder Abschnitt trägt eine Versal-Überschrift — vorher
                  hatten nur „Teilnahme" und „Veröffentlichung" eine, die vier
                  vorderen nicht. Auf einem Sub-Event-Reiter heißt der erste
                  Abschnitt wie der Termin selbst, nicht „Event". */}
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{scopeSub ? (childTermSingular || 'Sub-Event') : 'Event'}</div>
                <label className="dex-ui-label">
                  <StepBadge n={1} />
                  {!scopeSub && <span className="required">*</span>}
                  {scopeSub
                    ? (isDe ? `Wie heißt dieses ${childTermSingular || 'Sub-Event'}?` : `What is this ${childTermSingular || 'sub-event'} called?`)
                    : (isDe ? 'Wie heißt das Event?' : 'What is the event called?')}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> den offiziellen Namen des Events, z.B. <em>Sommerfest 2026</em> oder <em>JPMorgan Lauf 2026</em>.<br /><br />
                      <strong>Anzeige in der App:</strong> der Titel erscheint in der <strong>Eventliste</strong>, im <strong>Header der Detailseite</strong>, in <strong>Meine Events</strong> und im <strong>Admin Center</strong>. Wird auch für die Subsite-URL und für interne Verweise herangezogen.<br /><br />
                      <strong>Automatismen:</strong> der Titel wird 1:1 in den <strong>Betreff aller automatischen Mails</strong> übernommen (Anmelde-Bestätigung, Storno, Warteliste, Nachrück-Mail, QR-Code) sowie in den <strong>Outlook-Termin-Titel</strong> der Teilnehmer.<br /><br />
                      <strong>Empfehlung:</strong> sprechend wählen — der Titel ist das Erste, was Teilnehmer sehen, und identifiziert das Event in ihrem Outlook-Kalender.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the official name of the event, e.g. <em>Summer Party 2026</em> or <em>JPMorgan Run 2026</em>.<br /><br />
                      <strong>Shown in the app:</strong> shown in the <strong>event list</strong>, the <strong>detail page header</strong>, in <strong>My Events</strong> and the <strong>admin center</strong>. Also feeds the subsite URL and internal references.<br /><br />
                      <strong>Automation:</strong> the title is used 1:1 as the <strong>subject of every automated mail</strong> (registration, cancellation, waitlist, promotion, QR-code) and as the <strong>Outlook event title</strong> in attendees{'’'} calendars.<br /><br />
                      <strong>Tip:</strong> pick something descriptive — it is the first thing attendees see and identifies the event in their Outlook calendar.
                    </>
                  )} />
                </label>
                {/* v28.89: dasselbe Feld, je nach Ebene auf Hauptevent oder
                    Sub-Event gebunden. Die Pflicht-Markierung gilt nur fürs
                    Hauptevent — ein Sub-Event ohne Titel blockiert den Schritt
                    nicht (getStepErrors prüft weiterhin den Top-Level-Titel). */}
                <input
                  className="form-input"
                  value={scTitle}
                  onChange={e => setScTitle(e.target.value)}
                  placeholder={scopeSub ? t('create.subevents.title.placeholder') : (isDe ? 'z.B. Sommerfest 2026' : 'e.g. Summer Party 2026')}
                  style={scopeSub ? undefined : errorBorderStyle('title')}
                />
                {!scopeSub && fieldHasError('title') && <span style={{ color: 'var(--dex-red)', fontSize: '0.75rem' }}>{t('create.error.required')}</span>}
                {!scopeSub && (
                  <div className="dex-ui-help">
                    {isDe
                      ? 'Steht auf der Kachel, in Meine Events und wird Betreff aller automatischen Mails und Titel des Outlook-Termins.'
                      : 'Shown on the tile, in My Events, and used as the subject of every automatic mail and the Outlook invite title.'}
                  </div>
                )}
              </div>

              {/* v9.24: Event-Datum direkt nach Title — auto-fillt die Deadlines.
                  Vorher in Step 1, jetzt in Step 0 weil das fundamentale Info ist. */}
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'Zeitraum' : 'Dates'}</div>
                <div className="dex-ui-label">
                  <StepBadge n={2} />
                  {scopeSub
                    ? (isDe ? `Wann findet dieses ${childTermSingular || 'Sub-Event'} statt?` : `When does this ${childTermSingular || 'sub-event'} take place?`)
                    : (isDe ? 'Wann findet das Event statt?' : 'When does the event take place?')}
                  {!scopeSub && <span className="required">*</span>}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> Start- und Endzeit des Events (Datum + Uhrzeit, jeweils Berliner Zeit).<br /><br />
                      <strong>Anzeige in der App:</strong> die Werte erscheinen <strong>oben auf der Anmelde-Seite</strong>, in der <strong>Eventliste</strong>, in <strong>Meine Events</strong> und im Admin Center. Die App nutzt sie auch für <strong>Sortierung</strong> (nächste Events zuerst) und für die Logik <strong>Event ist vorbei</strong> (danach werden manche Aktionen wie Anmeldung gesperrt).<br /><br />
                      <strong>Automatismen:</strong> Datum + Uhrzeit landen 1:1 im <strong>Outlook-Termin der Teilnehmer</strong> — der Termin blockt damit den richtigen Slot im Kalender. Außerdem werden die <strong>Anmelde-Deadline</strong> (7 Tage vor Start) und die <strong>Letzte Abmeldemöglichkeit</strong> (3 Tage vor Start) automatisch vorgeschlagen — beide kannst du im Schritt <strong>Kapazität & Sichtbarkeit</strong> jederzeit überschreiben.<br /><br />
                      <strong>Auswirkung für Teilnehmer:</strong> sehen das Datum sofort in der Liste, bekommen es in jeder Bestätigungs-Mail und als Outlook-Eintrag.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> start and end (date + time, Berlin local time).<br /><br />
                      <strong>Shown in the app:</strong> shown <strong>at the top of the registration page</strong>, in the <strong>event list</strong>, in <strong>My Events</strong> and the admin center. Also drives <strong>sort order</strong> (upcoming first) and the <strong>event is over</strong> logic (some actions get locked after that).<br /><br />
                      <strong>Automation:</strong> date + time go 1:1 into the <strong>attendee{'’'}s Outlook event</strong> so the right slot gets blocked. The <strong>registration deadline</strong> (7 days before start) and <strong>last cancellation date</strong> (3 days before start) are auto-suggested — both can be overridden in step <strong>Capacity & Visibility</strong>.<br /><br />
                      <strong>Effect for attendees:</strong> they see the date in the list, in every confirmation email and as an Outlook entry.
                    </>
                  )} />
                </div>
                <p className="dex-ui-help" style={{ margin: '-2px 0 10px' }}>
                  {isDe
                    ? 'Berliner Zeit. Landet 1:1 im Outlook-Termin der Teilnehmer; Anmelde- und Abmeldefrist werden daraus vorgeschlagen.'
                    : 'Berlin time. Goes 1:1 into the attendees’ Outlook invite; registration and cancellation deadlines are suggested from it.'}
                </p>
              <div className="dex-ui-grid-2">
                <div className="dex-ui-field">
                  <label className="dex-ui-label">
                    {!scopeSub && <span className="required">*</span>} {isDe ? 'Beginn' : 'Start'}
                    <InfoTooltip text={isDe ? (
                      <>
                        <strong>Startzeitpunkt</strong> — Datum + Uhrzeit, ab wann das Event läuft. Wandert 1:1 in den <strong>Outlook-Termin</strong> jedes Teilnehmers (blockt den Kalender-Slot) und in <strong>jede Bestätigungs-Mail</strong>. Bestimmt außerdem die Standard-Vorschläge für <strong>Anmelde-Deadline</strong> (7 Tage vor Start) und <strong>Letzte Abmeldemöglichkeit</strong> (3 Tage vor Start).
                      </>
                    ) : (
                      <>
                        <strong>Start time</strong> — date + time when the event begins. Goes 1:1 into every attendee Outlook event (blocks the calendar slot) and into every <strong>confirmation email</strong>. Also drives the auto-suggestions for <strong>registration deadline</strong> (7 days before start) and <strong>last cancellation date</strong> (3 days before start).
                      </>
                    )} />
                  </label>
                  <DatePicker
                    selected={scStart}
                    onChange={setScStart}
                    // v29.52: Bei „ganztägig" gibt es nichts zu wählen — die
                    // Uhrzeit-Spalte stehen zu lassen lädt zum Widerspruch ein.
                    showTimeSelect={!scAllDay}
                    timeFormat="HH:mm"
                    timeIntervals={15}
                    timeCaption={isDe ? 'Uhrzeit' : 'Time'}
                    dateFormat={scAllDay ? 'dd.MM.yyyy' : 'dd.MM.yyyy, HH:mm'}
                    locale="de"
                    // v28.66: Beim Sub-Event heißt leer „Zeit des Hauptevents".
                    placeholderText={scopeSub ? t('create.subevents.time.placeholder') : (isDe ? 'Datum und Uhrzeit wählen' : 'Choose date and time')}
                    className="form-input"
                    wrapperClassName="dex-datepicker-wrapper"
                    calendarClassName="dex-datepicker-calendar"
                    popperPlacement="bottom-start"
                    maxDate={scopeSub ? (scEnd || undefined) : undefined}
                    isClearable
                    autoComplete="off"
                  />
                  {!scopeSub && fieldHasError('startDate') && <span style={{ color: 'var(--dex-red)', fontSize: '0.75rem' }}>{t('create.error.required')}</span>}
                </div>
                <div className="dex-ui-field">
                  <label className="dex-ui-label">
                    {!scopeSub && <span className="required">*</span>} {isDe ? 'Ende' : 'End'}
                    <InfoTooltip text={isDe ? (
                      <>
                        <strong>Endzeitpunkt</strong> — Datum + Uhrzeit, wann das Event vorbei ist. Wandert 1:1 in den <strong>Outlook-Termin</strong> der Teilnehmer (sonst läuft der Termin endlos). Wichtig auch für interne Logik: nach diesem Zeitpunkt zählt das Event als <strong>vorbei</strong> — Anmeldungen werden gesperrt, das Event rutscht in der Liste nach unten und manche automatische Benachrichtigungen (z.B. Late-Cancel-Hinweise) reagieren darauf.
                      </>
                    ) : (
                      <>
                        <strong>End time</strong> — date + time when the event finishes. Goes 1:1 into the attendee Outlook event (otherwise it would never end). Also feeds internal logic: past this point the event counts as <strong>over</strong> — registrations get locked, it drops down the list, and some automated notifications (e.g. late-cancel alerts) react to it.
                      </>
                    )} />
                  </label>
                  <DatePicker
                    selected={scEnd}
                    onChange={setScEnd}
                    // v29.52: Bei „ganztägig" gibt es nichts zu wählen — die
                    // Uhrzeit-Spalte stehen zu lassen lädt zum Widerspruch ein.
                    showTimeSelect={!scAllDay}
                    timeFormat="HH:mm"
                    timeIntervals={15}
                    timeCaption={isDe ? 'Uhrzeit' : 'Time'}
                    dateFormat={scAllDay ? 'dd.MM.yyyy' : 'dd.MM.yyyy, HH:mm'}
                    locale="de"
                    placeholderText={scopeSub ? t('create.subevents.time.placeholder') : (isDe ? 'Datum und Uhrzeit wählen' : 'Choose date and time')}
                    className="form-input"
                    wrapperClassName="dex-datepicker-wrapper"
                    calendarClassName="dex-datepicker-calendar"
                    popperPlacement="bottom-start"
                    minDate={scStart || undefined}
                    isClearable
                    autoComplete="off"
                  />
                  {!scopeSub && fieldHasError('endDate') && <span style={{ color: 'var(--dex-red)', fontSize: '0.75rem' }}>{t('create.error.required')}</span>}
                </div>
              </div>
              {!scopeSub && fieldHasError('endBeforeStart') && <p style={{ color: 'var(--dex-red)', fontSize: '0.8rem', marginTop: -4, marginBottom: 8 }}>{t('create.error.endBeforeStart')}</p>}
              {/* v28.89: Ende-vor-Start je Sub-Event — dieselbe Prüfung, die
                  bisher an der Sub-Event-Karte hing. */}
              {scopeSub && scStart && scEnd && (scAllDay ? dayKeyOfDate(scEnd) < dayKeyOfDate(scStart) : scEnd <= scStart) && (
                <p style={{ color: 'var(--dex-red, #c00)', fontSize: '0.8rem', marginTop: -4, marginBottom: 8 }}>
                  {isDe
                    ? 'Das Enddatum dieses Sub-Events liegt vor dem Startdatum — bitte korrigieren.'
                    : 'The end date of this sub-event is before the start date — please correct it.'}
                </p>
              )}
              {/* v29.52: Ganztägiger Termin. Bisher gab es dafür nur 00:00–23:59
                  — in Outlook ist das ein normaler Termin über den ganzen Tag,
                  der die Verfügbarkeit auf „gebucht" setzt, statt oben im
                  Kalenderkopf als Ganztags-Eintrag zu stehen. */}
              {/* v31.2: Die beiden Outlook-Fragen als Toggle-Zeilen nebeneinander
                  — jede mit der Folge in einer Zeile. */}
              <div className="dex-ui-grid-2" style={{ marginTop: 12 }}>
              <label className={cx('dex-ui-toggle-row', scAllDay && 'is-active')}>
                <input type="checkbox" checked={scAllDay} onChange={e => setScAllDay(e.target.checked)} />
                <span className="dex-ui-toggle-row-body">
                  <span className="dex-ui-toggle-row-title">{isDe ? 'Ganztägiger Termin' : 'All-day event'}</span>
                  <span className="dex-ui-toggle-row-desc">
                    {isDe
                      ? 'Steht in Outlook oben im Kalenderkopf statt als Block über den Tag und lässt die Verfügbarkeit frei. Ohne Haken bucht ein Termin von 00:00 bis 23:59 den ganzen Tag als belegt.'
                      : 'Appears in the Outlook calendar header instead of as a block across the day and leaves attendees shown as free. Without it, a 00:00–23:59 entry books the whole day as busy.'}
                  </span>
                </span>
              </label>
              {/* v29.54: Kalender blockieren ja/nein. Der Haken ist ANGEHAKT =
                  beschäftigt (Default, bisheriges Verhalten); gespeichert wird
                  der umgekehrte Wert `showAsFree` — siehe Kommentar an
                  DeloitteEvent.showAsFree. Die Umkehrung passiert genau hier,
                  an einer Stelle, und nirgends sonst. */}
              {/* v31.2: Sichtbar bleibt eine Zeile Folge; der Satz zur
                  Arbeitstag-Belegung bei „ganztägig" liegt im Tooltip am Titel
                  — sonst lief die Zeile auf drei bis vier Zeilen. */}
              <label className={cx('dex-ui-toggle-row', !scShowAsFree && 'is-active')}>
                <input type="checkbox" checked={!scShowAsFree} onChange={e => setScShowAsFree(!e.target.checked)} />
                <span className="dex-ui-toggle-row-body">
                  <span className="dex-ui-toggle-row-title">
                    {isDe ? 'Termin blockiert den Kalender' : 'Entry blocks the calendar'}
                    <InfoTooltip text={isDe ? (
                      <>
                        <strong>Mit Haken</strong> steht der Termin bei den Teilnehmern auf <strong>Beschäftigt</strong>, <strong>ohne Haken</strong> auf <strong>Frei</strong> — er blockiert dann nichts.<br /><br />
                        <strong>Bei einem ganztägigen Termin</strong> gilt mit Haken der <strong>komplette Arbeitstag</strong> als belegt; für ein Angebot, zu dem man nur zeitweise dazukommt, nimmst du den Haken besser raus.
                      </>
                    ) : (
                      <>
                        <strong>Ticked</strong>, the entry shows as <strong>busy</strong> for attendees; <strong>unticked</strong> it shows as <strong>free</strong> and blocks nothing.<br /><br />
                        <strong>For an all-day entry</strong> ticking marks the <strong>entire working day</strong> as taken; for something people only drop into, better untick it.
                      </>
                    )} />
                  </span>
                  <span className="dex-ui-toggle-row-desc">
                    {isDe
                      ? <>Steht bei den Teilnehmern auf <strong>Beschäftigt</strong>; ohne Haken als <strong>Frei</strong>.</>
                      : <>Shows as <strong>busy</strong> for attendees; unticked it shows as <strong>free</strong>.</>}
                  </span>
                </span>
              </label>
              </div>
              <p className="dex-ui-help" style={{ marginTop: 8 }}>
                {scAllDay
                  ? (isDe
                    ? 'Ganztägig: Es zählt nur das Datum — die Uhrzeit spielt für den Outlook-Termin keine Rolle mehr.'
                    : 'All-day: only the date matters — the time is no longer used for the Outlook entry.')
                  : scopeSub
                    ? (isDe
                      ? 'Leer lassen heißt: Dieses Sub-Event übernimmt die Zeiten des Hauptevents. Die Uhrzeit wird für den Outlook-Kalendereintrag der Teilnehmer verwendet.'
                      : 'Leaving these empty means the sub-event inherits the main event’s times. The time is used for the attendees’ Outlook entry.')
                    : (isDe
                      ? 'Die Uhrzeit wird für den Outlook-Kalendereintrag der Teilnehmer verwendet.'
                      : 'The time is used for the attendees’ Outlook calendar entry.')}
              </p>
              </div>

              <div className="dex-ui-section">
                {/* v28.7: kein <label> mehr, sondern <div> — rechts sitzt jetzt
                    der „Beschreibung anzeigen"-Schalter mit eigenem <label>;
                    verschachtelte Labels würden Klicks auf die Überschrift
                    fälschlich auf den Schalter umleiten. */}
                <div className="dex-ui-section-title">{isDe ? 'Beschreibung' : 'Description'}</div>
                <div className="dex-ui-label" style={{ flexWrap: 'wrap' }}>
                  <StepBadge n={3} />
                  {isDe ? 'Was sollen Teilnehmer vorab wissen?' : 'What should attendees know beforehand?'}
                  <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> die <strong>Hauptbeschreibung</strong> des Events — was findet statt, an wen richtet es sich, was sollten Teilnehmer wissen.<br /><br />
                      <strong>Anzeige in der App:</strong> wird oben auf der <strong>Anmelde-Seite</strong> und unter <strong>Meine Events</strong> angezeigt. Du kannst <strong>HTML-Formatierung</strong> nutzen (Fettdruck, Listen, Links, Bilder) — die App rendert sie 1:1.<br /><br />
                      <strong>Automatismen:</strong> wenn der Outlook-Termin-Body leer ist, wird die Beschreibung als <strong>Fallback in den Outlook-Kalendereintrag</strong> übernommen.<br /><br />
                      <strong>Auswirkung für Teilnehmer:</strong> erste inhaltliche Information vor der Anmeldung. Optional — leer ist erlaubt, aber bei externen Empfängerkreisen empfehlenswert.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>main event description</strong> — what is happening, who it is for, what attendees should know.<br /><br />
                      <strong>Shown in the app:</strong> shown at the top of the <strong>registration page</strong> and under <strong>My Events</strong>. You can use <strong>HTML formatting</strong> (bold, lists, links, images) — the app renders it 1:1.<br /><br />
                      <strong>Automation:</strong> if the Outlook event body is empty, the description is used as a <strong>fallback in the Outlook calendar entry</strong>.<br /><br />
                      <strong>Effect for attendees:</strong> first piece of substantive information before registering. Optional — leaving it blank is fine, but recommended for broader audiences.
                    </>
                  )} />
                  {/* v28.7: „Keine Beschreibung nutzen" direkt neben der
                      Überschrift (Default: Beschreibung nutzen). Anhaken
                      leert die Beschreibung und blendet den Editor aus. */}
                  {/* v28.89: „Keine Beschreibung" ist eine Entscheidung fürs
                      Hauptevent (sie steckt als Flag in EmailTemplateOverrides).
                      Ein Sub-Event lässt seine Beschreibung schlicht leer. */}
                  {/* v31.2: Schalter mit positiver Aussage („anzeigen") statt
                      Häkchen „Keine Beschreibung nutzen" — gespeichert wird
                      weiter `noDescription`, die Umkehr passiert nur hier.
                      2a′: direkt hinter der Beschriftung, nicht rechts außen. */}
                  {!scopeSub && (
                    <label className="dex-ui-switch" style={{ display: 'inline-flex', marginLeft: 12 }}>
                      <input
                        type="checkbox"
                        checked={!noDescription}
                        onChange={e => {
                          const on = !e.target.checked;
                          setNoDescription(on);
                          if (on) setDescription('');
                        }}
                      />
                      <span className="dex-ui-switch-track" />
                      <span className="dex-ui-switch-label" style={{ fontWeight: 500, fontSize: '0.8rem' }}>{isDe ? 'Beschreibung anzeigen' : 'Show a description'}</span>
                    </label>
                  )}
                </div>
                {/* v9.39: Beschreibung als HTML-Editor (vorher plain textarea).
                    Live-Vorschau im HtmlEditorModal — wird auf der Anmelde-Seite
                    1:1 als HTML gerendert.
                    v28.7: Die frühere Starthilfe-Box (Tipp-Text + Vorschläge,
                    v26.77) lebt jetzt IM Editor-Dialog (headerExtra +
                    bodyTemplates) — der Wizard-Schritt bleibt schlank. */}
                {(!scopeSub && noDescription) ? (
                  <div className="dex-ui-card dex-ui-card--soft dex-ui-card--muted dex-ui-muted">
                    {isDe ? 'Auf der Anmelde-Seite wird keine Beschreibung angezeigt. Schalter oben einschalten, um eine zu schreiben.' : 'No description will be shown on the registration page. Turn the switch above on to write one.'}
                  </div>
                ) : (
                // v31.2 (2a′): Auszug und Knopf stehen zusammen links (kein
                // flex:1, das den Knopf an den Rand schiebt); die Kachel selbst
                // öffnet den Editor — Hover, Zeiger, Enter/Leertaste.
                <div
                  className="dex-ui-card dex-ui-card--soft dex-ui-card--hover"
                  role="button"
                  tabIndex={0}
                  onClick={openDescriptionEditor}
                  onKeyDown={rowKeyHandler(openDescriptionEditor)}
                  style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap', cursor: 'pointer' }}
                >
                  <span className="dex-ui-muted" style={{ minWidth: 200, lineHeight: 1.5 }}>
                    {scDescription
                      ? `${scDescription.replace(/<[^>]+>/g, '').substring(0, 120)}${scDescription.length > 120 ? '…' : ''}`
                      : (isDe ? 'Noch keine Beschreibung — sie steht oben auf der Anmeldeseite und in Meine Events. HTML-Formatierung ist möglich.' : 'No description yet — it appears at the top of the registration page and in My Events. HTML formatting is possible.')}
                  </span>
                  <button
                    type="button"
                    className="btn btn-secondary dex-ui-btn-sm"
                    onClick={e => { e.stopPropagation(); openDescriptionEditor(); }}
                    onKeyDown={stopBubble}
                  >
                    {scDescription ? (isDe ? 'Bearbeiten & Vorschau' : 'Edit & preview') : (isDe ? 'Beschreibung schreiben' : 'Write a description')}
                  </button>
                </div>
                )}
                {/* v18.73: Hinweis, wenn Name/Datum/Ort des Events redundant in
                    der Beschreibung stehen — die werden bereits separat auf der
                    Anmelde-Seite angezeigt. Mit klickbarem Beispieltext. */}
                {(() => {
                  // v28.89: Der Redundanz-Hinweis vergleicht mit Name/Ort/Datum
                  // des Hauptevents — auf einem Sub-Event-Reiter wäre er
                  // irreführend.
                  if (scopeSub) return null;
                  const plain = (description || '')
                    .replace(/<[^>]+>/g, ' ')
                    .replace(/&nbsp;/gi, ' ')
                    .replace(/\s+/g, ' ')
                    .toLowerCase();
                  if (plain.trim().length < 8) return null;
                  const hits: string[] = [];
                  const tl = title.trim().toLowerCase();
                  if (tl.length >= 4 && plain.indexOf(tl) >= 0) hits.push(isDe ? 'der Event-Name' : 'the event name');
                  const locl = location.trim().toLowerCase();
                  if (locl.length >= 4 && plain.indexOf(locl) >= 0) hits.push(isDe ? 'der Ort' : 'the location');
                  if (startDate) {
                    const d = new Date(startDate);
                    if (!isNaN(d.getTime())) {
                      const dd = String(d.getDate()).padStart(2, '0');
                      const mm = String(d.getMonth() + 1).padStart(2, '0');
                      const yyyy = String(d.getFullYear());
                      const monthsDe = ['januar', 'februar', 'märz', 'april', 'mai', 'juni', 'juli', 'august', 'september', 'oktober', 'november', 'dezember'];
                      const monthsEn = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
                      const mn = (isDe ? monthsDe : monthsEn)[d.getMonth()];
                      const pats = [`${dd}.${mm}.${yyyy}`, `${dd}.${mm}.`, `${d.getDate()}. ${mn}`, `${d.getDate()}.${mn}`, `${d.getDate()} ${mn}`];
                      if (pats.some(p => plain.indexOf(p) >= 0)) hits.push(isDe ? 'das Datum' : 'the date');
                    }
                  }
                  if (hits.length === 0) return null;
                  const joined = hits.length === 1
                    ? hits[0]
                    : hits.slice(0, -1).join(', ') + (isDe ? ' und ' : ' and ') + hits[hits.length - 1];
                  return (
                    <WizardHint
                      isDe={isDe}
                      title={isDe ? 'Beschreibung wiederholt Basis-Infos' : 'Description repeats basic info'}
                      style={{ marginTop: 12 }}
                    >
                      {isDe
                        ? <>In der Beschreibung steht offenbar <strong>{joined}</strong>. <strong>Name, Datum und Ort</strong> des Events werden bereits <strong>separat</strong> auf der Anmelde-Seite angezeigt — du musst sie hier nicht wiederholen. Nutze die Beschreibung lieber für einen einladenden, inhaltlichen Text. Über <strong>„Bearbeiten &amp; Vorschau“</strong> kannst du einen Beispieltext übernehmen.</>
                        : <>Your description appears to contain <strong>{joined}</strong>. The event&rsquo;s <strong>name, date and location</strong> are already shown <strong>separately</strong> on the registration page — no need to repeat them here. Use the description for an inviting, substantive text instead. Via <strong>“Edit &amp; Preview”</strong> you can load an example text.</>}
                    </WizardHint>
                  );
                })()}
              </div>

              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'Bild' : 'Image'}</div>
                <div className="dex-ui-label">
                  <StepBadge n={4} />
                  {scopeSub
                    ? (isDe ? `Welches Bild zeigt dieses ${childTermSingular || 'Sub-Event'}?` : `Which image shows this ${childTermSingular || 'sub-event'}?`)
                    : (isDe ? 'Welches Bild zeigt das Event?' : 'Which image represents the event?')}
                  <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> ein <strong>Hauptbild fürs Event</strong> (Foto vom Veranstaltungsort, Eventlogo, Stimmungsbild). Wird zentral als Item-Attachment am Event gespeichert.<br /><br />
                      <strong>Anzeige in der App:</strong> erscheint <strong>oben auf der Anmelde-Seite</strong>, <strong>als Kachel-Hintergrund</strong> in der Eventliste und <strong>in Meine Events</strong>. Macht Events visuell unterscheidbar in einer langen Liste.<br /><br />
                      <strong>Empfehlung:</strong> Querformat (z.B. 16:9), gute Auflösung (mind. 1200px breit). Hochformat funktioniert auch — die App erkennt das automatisch und legt das Bild dann links neben den Detail-Rows ab statt als Banner.<br /><br />
                      <strong>Auswirkung für Teilnehmer:</strong> visueller Wiedererkennungswert in ihrer Eventliste und auf der Anmelde-Seite.<br /><br />
                      <strong>Nicht dasselbe wie das Mail-Logo:</strong> Mails und Outlook-Termin zeigen das Logo aus dem Schritt <strong>Kommunikation</strong>, nicht dieses Bild. Lässt du hier leer, greift die Anmelde-Seite seit v29.13 auf das Mail-Logo zurück — damit dort nicht der generische DEX-Kreis steht.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>main event image</strong> (venue photo, event logo, mood shot). Stored centrally as an item attachment.<br /><br />
                      <strong>Shown in the app:</strong> shown <strong>at the top of the registration page</strong>, <strong>as the tile background</strong> in the event list and <strong>under My Events</strong>. Makes events visually distinguishable in a long list.<br /><br />
                      <strong>Tip:</strong> landscape (e.g. 16:9), high resolution (at least 1200px wide). Portrait works too — the app detects orientation and places the image to the left of the detail rows instead of as a banner.<br /><br />
                      <strong>Effect for attendees:</strong> visual recognition in their list and on the registration page.<br /><br />
                      <strong>Not the same as the mail logo:</strong> emails and the Outlook invite use the logo from the <strong>Communication</strong> step, not this image. If you leave this empty, the registration page falls back to the mail logo (since v29.13) instead of showing the generic DEX circle.
                    </>
                  )} />
                </div>
                <p className="dex-ui-help" style={{ margin: '-2px 0 10px' }}>
                  {scopeSub
                    ? (isDe ? 'Vorschaubild neben diesem Termin auf der Anmeldeseite.' : 'Thumbnail next to this date on the registration page.')
                    : (isDe ? 'Kachel in der Event-Liste und Kopf der Anmeldeseite. Mails und Outlook-Termin haben ihr eigenes Logo (Schritt Kommunikation).' : 'Tile in the event list and top of the registration page. Emails and the Outlook invite use their own logo (Communication step).')}
                </p>
                {/* v31.2: Aktionen als Knopfreihe UNTER dem Bild statt als
                    Overlays ohne Hover — drei halbtransparente Flächen auf
                    einem Foto lasen sich als Bildteil, nicht als Knöpfe. */}
                {scImagePreview && (
                  <div className="dex-ui-card dex-ui-card--soft" style={{ padding: 12, marginBottom: 10 }}>
                    <img
                      src={scImagePreview}
                      alt="Vorschau"
                      style={{
                        // Korrekte Auflösung beibehalten, nur in der Höhe begrenzen + max-Breite zur Sicherheit
                        display: 'block',
                        maxHeight: 220,
                        maxWidth: '100%',
                        width: 'auto',
                        height: 'auto',
                        objectFit: 'contain',
                        borderRadius: 'var(--dex-radius)',
                        background: 'var(--dex-gray-100)',
                      }}
                    />
                    <div className="dex-ui-inline" style={{ marginTop: 10 }}>
                    {/* v23.15: Bild editieren (zuschneiden / auf Kreis). */}
                    <button
                      type="button"
                      className="btn btn-secondary dex-ui-btn-sm"
                      onClick={() => {
                        // v28.89: Sub-Event-Bilder haben ihr eigenes
                        // Zuschnitt-Modal (subImageCropIdx) — Ziel ist der
                        // gerade gewählte Reiter.
                        if (scopeSub) { setSubImageCropIdx(activeScopeIdx - 1); return; }
                        // v28.12: Beim Editieren eines BESTEHENDEN Bilds das
                        // Original einfangen, BEVOR der Zuschnitt Preview/File
                        // ersetzt — sonst kann die Event-Liste nach einem
                        // Kreis-Zuschnitt nicht aufs Querformat-Original
                        // zurückfallen (frischer Upload fängt es im
                        // onChange-Handler ein, dieser Pfad hier fehlte).
                        if (!imageOrigFile && imagePreview) {
                          void (async () => {
                            try {
                              const resp = await fetch(imagePreview, imagePreview.indexOf('data:') === 0 ? undefined : { credentials: 'include' });
                              const blob = await resp.blob();
                              const f = new File([blob], `event-image-orig.${(blob.type || '').indexOf('png') >= 0 ? 'png' : 'jpg'}`, { type: blob.type || 'image/jpeg' });
                              const objUrl = URL.createObjectURL(blob);
                              const probe = new Image();
                              probe.onload = () => {
                                URL.revokeObjectURL(objUrl);
                                if (probe.naturalHeight > 0) {
                                  const r = probe.naturalWidth / probe.naturalHeight;
                                  // Nur Querformat-Quellen taugen als Original —
                                  // ein bereits runder/quadratischer Bestand
                                  // bleibt unangetastet (gespeichertes Original
                                  // wird dann NICHT überschrieben/gelöscht).
                                  if (r >= 1.2) { setImageOrigFile(f); setImageOrigAspect(r); }
                                }
                              };
                              probe.onerror = () => URL.revokeObjectURL(objUrl);
                              probe.src = objUrl;
                            } catch { /* best-effort */ }
                          })();
                        }
                        setImageEditOpen(true);
                      }}
                    >
                      <Icon iconName="Crop" style={{ fontSize: 13 }} /> {isDe ? 'Zuschneiden' : 'Crop'}
                    </button>
                    {/* v30.98: Bild herunterladen (Nutzer-Ansage 07.09.2026: „ich
                        möchte auch die Möglichkeit haben, das Bild zu speichern").
                        Data-URL direkt, http-URL über fetch → Blob, damit der
                        Browser speichert statt das Bild nur zu öffnen. */}
                    <button
                      type="button"
                      className="btn btn-secondary dex-ui-btn-sm"
                      onClick={() => {
                        void (async () => {
                          try {
                            const src = scImagePreview;
                            const isData = src.indexOf('data:') === 0;
                            const blob = isData ? await (await fetch(src)).blob() : await (await fetch(src, { credentials: 'include' })).blob();
                            const ext = (blob.type || '').indexOf('png') >= 0 ? 'png' : (blob.type || '').indexOf('webp') >= 0 ? 'webp' : 'jpg';
                            const url = URL.createObjectURL(blob);
                            const a = document.createElement('a');
                            a.href = url; a.download = `${(title || 'Event-Bild').replace(/[^\w\-äöüÄÖÜß ]+/g, '').trim().replace(/\s+/g, '_') || 'Event-Bild'}.${ext}`; a.style.display = 'none';
                            document.body.appendChild(a); a.click();
                            setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 400);
                          } catch { window.open(scImagePreview, '_blank'); }
                        })();
                      }}
                      title={isDe ? 'Bild als Datei speichern' : 'Save image as file'}
                    >
                      <Icon iconName="Download" style={{ fontSize: 13 }} /> {isDe ? 'Speichern' : 'Save'}
                    </button>
                    <button
                      type="button"
                      className="dex-ui-textbtn dex-ui-textbtn--danger"
                      style={{ marginLeft: 'auto' }}
                      onClick={() => {
                        // v28.89: Auf einem Sub-Event-Reiter betrifft das
                        // Entfernen dessen Bild — imageRemoved ist das Signal
                        // für den Save, das gespeicherte Attachment zu löschen.
                        if (scopeSub) { patchScopeSub({ imagePreview: '', imageFile: null, imageRemoved: true }); return; }
                        setImageFile(null); setImagePreview(''); setEventImageUrl(''); setImageOrigFile(null); setImageOrigAspect(null);
                      }}
                    >
                      <X size={14} /> {isDe ? 'Bild entfernen' : 'Remove image'}
                    </button>
                    </div>
                  </div>
                )}
                {/* v23.15: Bild-Zuschnitt-Modal — liefert das Ergebnis als
                    Data-URL (Vorschau) + File (Upload) zurück.
                    v23.25: „Darstellung pro Ansicht" lebt jetzt IN diesem Modal
                    (children), damit alle Bild-Einstellungen an einem Ort sind. */}
                {/* v26.97: Zuschneiden des Mail-/Outlook-Kopfbildes — gleiches
                    Modal, Ziel via logoCropTarget. Ergebnis (Data-URL) direkt in
                    das jeweilige Logo. */}
                {logoCropTarget && (
                  <ImageCropModal
                    open={!!logoCropTarget}
                    src={logoCropTarget === 'email' ? emailLogoPreview : outlookLogoPreview}
                    isDe={isDe}
                    allowAspect
                    onClose={() => setLogoCropTarget(null)}
                    onApply={(dataUrl) => {
                      // v28.10: Crop-Ergebnis ggf. verkleinern (2-MB-Schutz).
                      void shrinkLogoB64(dataUrl).then(small => {
                        if (logoCropTarget === 'email') setEmailLogoPreview(small);
                        else setOutlookLogoPreview(small);
                      });
                      setLogoCropTarget(null);
                    }}
                  />
                )}
                <ImageCropModal
                  open={imageEditOpen}
                  src={imagePreview}
                  isDe={isDe}
                  recommendCircle
                  onClose={() => setImageEditOpen(false)}
                  onApply={async (dataUrl, file) => {
                    setImagePreview(dataUrl);
                    setImageFile(file);
                    setImageEditOpen(false);
                    // v28.32: Der Kreis-Zuschnitt gilt nur noch für die
                    // Anmeldeseite und die Event-Karte. Mail-Kopf und
                    // Outlook-Termin sind RECHTECKIG — dort steht ab jetzt
                    // automatisch das UNBESCHNITTENE Original (v26.95 fragte
                    // stattdessen nach und nahm dann den Kreis, der im
                    // rechteckigen Kopf sichtbar angeschnitten ankam).
                    // Ein bereits bewusst gesetztes eigenes Kopfbild wird nicht
                    // überschrieben; wer den Kreis doch im Kopf haben will,
                    // nimmt „Bild auswählen".
                    try {
                      let srcFile: File | null = imageOrigFile;
                      if (!srcFile && editEvent && editEvent.imageOrigUrl) {
                        try {
                          const resp = await fetch(editEvent.imageOrigUrl, { credentials: 'include' });
                          const blob = await resp.blob();
                          srcFile = new File([blob], 'event-photo.jpg', { type: blob.type || 'image/jpeg' });
                        } catch { /* Original nicht ladbar → Zuschnitt nehmen */ }
                      }
                      const b64 = await fileToBase64(await compressImage(srcFile || file, 600, 0.85, true));
                      if (b64) {
                        if (!emailLogoPreview) { setEmailLogoPreview(b64); setEmailLogoFromPhoto(true); }
                        if (!outlookLogoPreview) { setOutlookLogoPreview(b64); setOutlookLogoFromPhoto(true); }
                      }
                    } catch { /* Kopfbild ist optional — Fehler nie durchreichen */ }
                  }}
                >
                  {/* v23.19/v23.25: Optional & einklappbar — Bild pro Ansicht
                      anders zoomen/skalieren. Default zu; wer einfach nur ein
                      Foto hochlädt, muss hier nichts tun. */}
                  <div>
                    <button
                      type="button"
                      className={cx('dex-ui-disclosure', imageDisplayOpen && 'is-open')}
                      onClick={() => setImageDisplayOpen(o => !o)}
                    >
                      <span className="dex-ui-disclosure-chevron"><Icon iconName="ChevronRight" style={{ fontSize: 12 }} /></span>
                      {isDe ? 'Darstellung pro Ansicht anpassen' : 'Adjust display per view'}
                      <span className="dex-ui-disclosure-count">{isDe ? 'optional' : 'optional'}</span>
                    </button>
                    {imageDisplayOpen && (
                      <div className="dex-ui-disclosure-body">
                        <p className="dex-ui-help" style={{ margin: '0 0 12px' }}>
                          {isDe
                            ? 'Zoom und Größe je Ansicht getrennt einstellen, damit das Bild überall gut sitzt. Nichts einstellen = das Bild füllt den Bereich zentriert.'
                            : 'Set zoom and size per view so the image sits well everywhere. Leave untouched = the image fills the area centered.'}
                        </p>
                        {([
                          { key: 'card' as const, label: isDe ? 'Event-Liste / Karte' : 'Event list / card', w: 240, h: 135 },
                          { key: 'hero' as const, label: isDe ? 'Anmeldeseite (Bild oben)' : 'Registration page (top image)', w: 200, h: 200 },
                        ]).map(view => {
                          const v: ImgView = imageDisplay[view.key] || { zoom: 1, posY: 50 };
                          const setV = (next: ImgView): void => setImageDisplay(prev => ({ ...prev, [view.key]: next }));
                          const isHero = view.key === 'hero';
                          const heroH = v.height ?? 340;
                          if (isHero) {
                            // v23.24: Anmeldeseite-Vorschau 1:1 wie die echte
                            // Registrierungsseite rendern — weiße Hülle volle
                            // Breite, Bild „contain" mit der eingestellten max.
                            // Höhe + Zoom (identische Style-Logik wie
                            // RegistrationPage Hero). So sieht der Organizer die
                            // tatsächliche Größe, nicht nur eine Mini-Annäherung.
                            return (
                              <div key={view.key} style={{ marginBottom: 16 }}>
                                <div style={{ fontSize: '0.8rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 6 }}>{view.label}</div>
                                <div style={{ width: '100%', background: '#fff', borderRadius: 'var(--dex-radius)', overflow: 'hidden', display: 'flex', justifyContent: 'center', boxShadow: 'inset 0 0 0 1px var(--dex-gray-200)', padding: 4 }}>
                                  <img
                                    src={imagePreview}
                                    alt={view.label}
                                    style={{ display: 'block', margin: '0 auto', maxWidth: '100%', maxHeight: heroH, width: 'auto', height: 'auto', objectFit: 'contain', transform: `scale(${v.zoom})`, transformOrigin: 'center center' }}
                                  />
                                </div>
                                <div style={{ marginTop: 8 }}>
                                  <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)' }}>{isDe ? 'Größe (max. Höhe)' : 'Size (max. height)'}</label>
                                  <input type="range" min={140} max={500} step={5} value={heroH} onChange={e => setV({ ...v, height: parseInt(e.target.value, 10) })} style={{ width: '100%' }} />
                                  <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)' }}>{isDe ? 'Zoom' : 'Zoom'}</label>
                                  <input type="range" min={0.3} max={3} step={0.01} value={v.zoom} onChange={e => setV({ ...v, zoom: parseFloat(e.target.value) })} style={{ width: '100%' }} />
                                  <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" style={{ marginTop: 4 }} onClick={() => setImageDisplay(prev => { const n = { ...prev }; delete n[view.key]; return n; })}>
                                    {isDe ? 'Zurücksetzen' : 'Reset'}
                                  </button>
                                </div>
                              </div>
                            );
                          }
                          return (
                            <div key={view.key} style={{ display: 'flex', gap: 14, alignItems: 'flex-start', marginBottom: 14, flexWrap: 'wrap' }}>
                              <div style={{ width: view.w, height: view.h, flexShrink: 0, overflow: 'hidden', borderRadius: 6, background: '#fff', position: 'relative', boxShadow: 'inset 0 0 0 1px var(--dex-gray-200)' }}>
                                {/* Event-Liste/Karte: volles Bild (contain, kein
                                    Crop — der Kreis wird nie abgeschnitten),
                                    Größe per Skalierung. */}
                                <img
                                  src={imagePreview}
                                  alt={view.label}
                                  style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', objectFit: 'contain', transform: `scale(${v.zoom})`, transformOrigin: 'center center' }}
                                />
                              </div>
                              <div style={{ flex: 1, minWidth: 160 }}>
                                <div style={{ fontSize: '0.8rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 6 }}>{view.label}</div>
                                <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)' }}>{isDe ? 'Größe (kleiner = mehr weißer Rand)' : 'Size (smaller = more white margin)'}</label>
                                <input type="range" min={0.3} max={1.5} step={0.01} value={v.zoom} onChange={e => setV({ ...v, zoom: parseFloat(e.target.value) })} style={{ width: '100%' }} />
                                <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" style={{ marginTop: 4 }} onClick={() => setImageDisplay(prev => { const n = { ...prev }; delete n[view.key]; return n; })}>
                                  {isDe ? 'Zurücksetzen' : 'Reset'}
                                </button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                </ImageCropModal>
                {/* v31.2: Upload als gestrichelte Auswahl-Kachel (Klasse
                    dex-ui-choice liefert den Hover, den ein Inline-Style nicht
                    kann). Ohne Bild groß und einladend, mit Bild eine schmale
                    Zeile „Anderes Bild wählen". */}
                <label className="dex-ui-choice" style={{ borderStyle: 'dashed', alignItems: 'center', justifyContent: scImagePreview ? 'flex-start' : 'center', padding: scImagePreview ? '10px 14px' : '22px 16px', fontSize: '0.86rem', color: 'var(--dex-gray-600)', fontWeight: 600 }}>
                  <Plus size={16} />
                  {(scopeSub ? scopeSub.imageFile : imageFile)?.name || (scImagePreview
                    ? (isDe ? 'Anderes Bild wählen' : 'Choose a different image')
                    : (isDe ? 'Bild auswählen — bestmöglich Querformat, mind. 1200 px breit' : 'Choose an image — landscape if possible, at least 1200 px wide'))}
                  <input
                    type="file"
                    accept="image/*"
                    style={{ display: 'none' }}
                    onChange={e => {
                      const file = e.target.files && e.target.files[0];
                      if (!file) return;
                      // v28.89: Auf einem Sub-Event-Reiter landet das Bild im
                      // Draft dieses Sub-Events (imageRemoved zurücksetzen,
                      // sonst löscht der Save das gerade Hochgeladene wieder).
                      if (scopeSub) {
                        setImageUploadError('');
                        const readerSe = new FileReader();
                        readerSe.onload = ev => {
                          patchScopeSub({
                            imageFile: file,
                            imagePreview: (ev.target?.result as string) || '',
                            imageRemoved: false,
                          });
                        };
                        readerSe.readAsDataURL(file);
                        return;
                      }
                      if (file) {
                        setImageUploadError('');
                        setImageFile(file);
                        const reader = new FileReader();
                        reader.onload = ev => {
                          const dataUrl = ev.target?.result as string || '';
                          setImagePreview(dataUrl);
                          // v28.11: Original + Seitenverhältnis merken, BEVOR
                          // der Zuschnitt Preview/File ersetzt.
                          setImageOrigFile(file);
                          setImageOrigAspect(null);
                          const probe = new Image();
                          probe.onload = () => { if (probe.naturalHeight > 0) setImageOrigAspect(probe.naturalWidth / probe.naturalHeight); };
                          probe.src = dataUrl;
                          // v28.10: Direkt nach dem Upload den Zuschnitt-
                          // Dialog öffnen (mit Kreis-Empfehlung) — vorher
                          // musste man „Bild editieren" extra anklicken.
                          setImageEditOpen(true);
                        };
                        reader.readAsDataURL(file);
                      }
                    }}
                  />
                </label>
                {/* v28.89: Auf einem Sub-Event-Reiter gilt derselbe Upload für
                    dieses Sub-Event. Sein Bild erscheint als Vorschaubild neben
                    der Auswahl auf der Anmeldeseite (v27.11); Mails und
                    Outlook-Termin haben davon unabhängig ihr eigenes Kopfbild. */}
                {scopeSub && (
                  <p className="dex-ui-help">
                    {isDe
                      ? 'Erscheint als Vorschaubild neben diesem Sub-Event in der Auswahl auf der Anmeldeseite. Ohne eigenes Bild bleibt die Zeile dort schlicht ohne Vorschau.'
                      : 'Shown as a thumbnail next to this sub-event in the selection on the registration page. Without its own image the row simply has no preview.'}
                  </p>
                )}
                {imageUploadError && (
                  <p style={{ color: 'var(--dex-red, #c00)', fontSize: '0.8rem', marginTop: 4 }}>{imageUploadError}</p>
                )}
                {/* v28.5: Layout-Wahl fürs Event-Bild auf der Anmeldeseite —
                    Banner in voller Breite ÜBER den Infos (gut für breite
                    Querformat-Fotos) vs. kompakt links neben den Infos.
                    v28.10: nur noch bei Querformat-Bildern (Ratio >= 1.2)
                    anbieten — für Kreis-/Quadrat-/Hochkant-Bilder ergibt
                    das Banner-Layout keinen Sinn.
                    v31.2: direkt unter dem Bild, zu dem sie gehört — als
                    Toggle-Zeile mit der Folge in einer Zeile. */}
                {!scopeSub && (imagePreview || imageFile) && wizardImgAspect != null && wizardImgAspect >= 1.2 && (
                  <label className={cx('dex-ui-toggle-row', imageBanner && 'is-active')} style={{ marginTop: 12 }}>
                    <input type="checkbox" checked={imageBanner} onChange={e => setImageBanner(e.target.checked)} />
                    <span className="dex-ui-toggle-row-body">
                      <span className="dex-ui-toggle-row-title">{isDe ? 'Bild als Banner in voller Breite über den Event-Infos zeigen' : 'Show the image as a full-width banner above the event info'}</span>
                      <span className="dex-ui-toggle-row-desc">
                        {isDe
                          ? 'Empfohlen für breite Querformat-Fotos: Bild oben über die ganze Karte, Titel/Datum/Ort darunter. Aus = Bild sitzt kompakt links neben den Infos (Standard).'
                          : 'Recommended for wide landscape photos: image across the full card at the top, title/date/location below. Off = compact image to the left of the info (default).'}
                      </span>
                    </span>
                  </label>
                )}
              </div>

              {/* v28.90: Pflichtanmeldung — eine Einstellung DIESES
                  Sub-Events, deshalb hier bei seinen Grundlagen und nicht
                  (neunmal wiederholt) in der Liste auf der Klammer-Ebene.
                  v28.77: Der Haken wurde als „dieses Sub-Event ist buchbar"
                  missverstanden und darum bei ALLEN gesetzt — das Ergebnis
                  ist das Gegenteil einer Auswahl. Diesen Zustand benennen,
                  sobald er eintritt, mit einem Klick zum Zurücknehmen.
                  v31.2: eigener Abschnitt statt Anhang ans Bild — die Frage
                  hat mit dem Bild nichts zu tun. */}
              {scopeSub && (
                <div className="dex-ui-section">
                  <div className="dex-ui-section-title">{isDe ? 'Teilnahme' : 'Participation'}</div>
                  <label className={cx('dex-ui-toggle-row', !!scopeSub.mandatory && 'is-active')}>
                    <input type="checkbox" checked={!!scopeSub.mandatory} onChange={e => patchScopeSub({ mandatory: e.target.checked })} />
                    <span className="dex-ui-toggle-row-body">
                      <span className="dex-ui-toggle-row-title">{isDe ? `Jeder Teilnehmer muss dieses ${childTermSingular || 'Sub-Event'} mitbuchen (Pflicht)` : `Every attendee must book this ${childTermSingular || 'sub-event'} (mandatory)`}</span>
                      <span className="dex-ui-toggle-row-desc">
                        {isDe
                          ? <>Eine Anmeldung ohne diesen Termin ist dann nicht möglich. <strong>Nur setzen, wenn er für wirklich alle verpflichtend ist</strong> (z.B. eine Auftaktveranstaltung). Für &bdquo;darf gebucht werden&ldquo; ist der Haken nicht nötig.</>
                          : <>Registering without this date is then not possible. <strong>Only set this if it is truly compulsory for everyone</strong> (e.g. a kick-off). For &bdquo;may be booked&ldquo; the checkbox is not needed.</>}
                      </span>
                    </span>
                  </label>
                  {(() => {
                    const named = subEvents.filter(s => (s.title || '').trim());
                    const mandatoryCount = named.filter(s => s.mandatory).length;
                    if (named.length < 2 || mandatoryCount !== named.length || !scopeSub.mandatory) return null;
                    return (
                      <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginTop: 10, display: 'block' }}>
                        <strong>{isDe ? `Alle ${named.length} Sub-Events sind als Pflicht markiert` : `All ${named.length} sub-events are marked mandatory`}</strong>
                        <div style={{ marginTop: 3 }}>
                          {isDe
                            ? <>Damit gibt es faktisch <strong>keine Auswahl mehr</strong> — wer teilnehmen möchte, muss <strong>alle {named.length}</strong> mitbuchen; wer auch nur einen Termin nicht kann, kann sich gar nicht anmelden.{subEventsOnlyMode ? <> Bei diesem Klammerevent läuft die Anmeldung ohnehin ausschließlich über die Sub-Events — der Haken ist dafür <strong>nicht nötig</strong>.</> : null} Gemeint war vermutlich, dass die Sub-Events buchbar sind — dafür lässt du den Haken einfach weg.</>
                            : <>That leaves <strong>no choice at all</strong> — attendees must book <strong>all {named.length}</strong>; anyone unavailable for a single date cannot register.{subEventsOnlyMode ? <> For this bracket event registration runs via the sub-events anyway — the checkbox is <strong>not needed</strong> for that.</> : null}</>}
                        </div>
                        <button
                          type="button"
                          className="btn btn-secondary dex-ui-btn-sm"
                          style={{ marginTop: 8 }}
                          onClick={() => setSubEvents(prev => prev.map(s => ({ ...s, mandatory: false })))}
                        >
                          {isDe ? 'Pflicht bei allen entfernen' : 'Remove mandatory from all'}
                        </button>
                      </div>
                    );
                  })()}
                </div>
              )}

              {/* v9.21: Entwurf-Flag — Default an, damit der Organizer die
                  Test-Strecke in Ruhe aufbauen und das Test-Team durchspielen
                  lassen kann. v22.27: volle Breite.
                  v31.2: vom Anfang ans Ende des Schritts gewandert — erst
                  beschreibt man das Event, dann entscheidet man, wann es
                  jemand sieht. Bleibt event-weit (nur auf der Klammer-Ebene);
                  Badge 5 — die Nummern folgen der Reihenfolge (07.09.2026). */}
              {activeScopeIdx === 0 && (
                <div className="dex-ui-section">
                  <div className="dex-ui-section-title">{isDe ? 'Veröffentlichung' : 'Publishing'}</div>
                  <label className={cx('dex-ui-toggle-row', isFictive && 'is-active')} style={isFictive ? { borderColor: 'var(--dex-orange, #ed8b00)', background: 'rgba(237,139,0,0.06)' } : undefined}>
                    <input type="checkbox" checked={isFictive} onChange={e => setIsFictive(e.target.checked)} />
                    <span className="dex-ui-toggle-row-body">
                      <span className="dex-ui-toggle-row-title">
                        <StepBadge n={5} />
                        {isDe ? 'Noch nicht veröffentlichen — als Entwurf speichern' : 'Don’t publish yet — save as draft'}
                        <InfoTooltip text={t('create.fictive.hint')} />
                      </span>
                      {/* v31.2: Nur der erste Satz bleibt sichtbar — wie man live
                          geht, steht im Tooltip und unter „Ab wann … live gehen?". */}
                      <span className="dex-ui-toggle-row-desc">
                        {isDe
                          ? <>Standard. Nur Admins, Organizer und das Test-Team sehen das Event und können sich anmelden — niemand meldet sich versehentlich an.</>
                          : <>Default. Only admins, organizers and the test team see the event and can register — nobody signs up by accident.</>}
                      </span>
                    </span>
                  </label>
                  {/* v9.21: ActiveFrom direkt unter dem Entwurfs-Toggle — wenn
                      der Organizer ein Live-Datum setzt, geht das Event ab dann
                      auch wenn das Entwurf-Häkchen noch on ist. Optional. */}
                  <div className="dex-ui-field" style={{ marginTop: 14 }}>
                    <label className="dex-ui-label">
                      {isDe ? 'Ab wann soll das Event automatisch live gehen?' : 'When should the event go live automatically?'}
                      <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
                      <InfoTooltip text={isDe ? (
                        <>
                          <strong>Was du hier einstellst:</strong> einen Zeitpunkt, ab dem das Event automatisch live geht — auch wenn der <strong>Entwurf-Haken</strong> noch gesetzt ist.<br /><br />
                          <strong>Anzeige in der App:</strong> bis zu diesem Zeitpunkt sehen <strong>nur Admins, Organizer und Test-Team</strong> das Event. Ab dem gesetzten Datum prüft die App bei jedem Aufruf, ob die Zeit schon erreicht ist; falls ja, wird das Event in der allgemeinen Eventliste eingeblendet.<br /><br />
                          <strong>Auswirkung für Teilnehmer:</strong> bis zum Aktiv-ab-Zeitpunkt taucht das Event nicht in der Liste auf, kann nicht aufgerufen werden und bekommt keine Mails. Ab dem Stichtag ist es ganz normal anmeldbar.<br /><br />
                          <strong>Leer lassen</strong> = kein Auto-Go-Live. Du musst dann manuell den Entwurf-Haken entfernen oder im Admin Center auf <strong>Event aktivieren</strong> klicken.
                        </>
                      ) : (
                        <>
                          <strong>What you set here:</strong> a date/time at which the event automatically goes live — even if the <strong>draft toggle</strong> is still on.<br /><br />
                          <strong>Shown in the app:</strong> until that point, only <strong>admins, organizers and the test team</strong> see the event. Once the timestamp is reached, the app reveals the event in the general event list.<br /><br />
                          <strong>Effect for attendees:</strong> until the active-from date the event is not listed, not openable, and produces no mails. After the timestamp it behaves like any other published event.<br /><br />
                          <strong>Leave empty</strong> = no auto-go-live. Publish manually by clearing the draft toggle or by clicking <strong>Activate event</strong> in the admin center.
                        </>
                      )} />
                    </label>
                    <DatePicker
                      selected={activeFrom ? new Date(activeFrom) : null}
                      onChange={(date: Date | null) => setActiveFrom(date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}` : '')}
                      showTimeSelect
                      timeFormat="HH:mm"
                      timeIntervals={15}
                      timeCaption={isDe ? 'Uhrzeit' : 'Time'}
                      dateFormat="dd.MM.yyyy, HH:mm"
                      locale="de"
                      placeholderText={isDe ? 'Datum und Uhrzeit wählen' : 'Choose date and time'}
                      className="form-input"
                      wrapperClassName="dex-datepicker-wrapper"
                      calendarClassName="dex-datepicker-calendar"
                      popperPlacement="bottom-start"
                      isClearable
                      autoComplete="off"
                    />
                    <div className="dex-ui-help">
                      {isDe
                        ? 'Bis dahin sehen nur Admins, Organizer und Test-Team das Event. Leer = kein automatisches Go-Live, du aktivierst von Hand.'
                        : 'Until then only admins, organizers and the test team see the event. Empty = no automatic go-live; you activate it manually.'}
                    </div>
                  </div>
                  {/* v23.14: Vorschau-Wahl — nur sinnvoll bei gesetztem „Aktiv ab".
                      v31.2: zwei Kacheln statt zweier Radios — jede nennt die Folge. */}
                  {activeFrom && (
                    <div className="dex-ui-field">
                      <div className="dex-ui-label">{isDe ? 'Bis zum Aktivierungszeitpunkt ist das Event …' : 'Until the activation time the event is …'}</div>
                      <div className="dex-ui-grid-2" role="radiogroup">
                        <button type="button" role="radio" aria-checked={!previewBeforeActive} className={cx('dex-ui-choice', !previewBeforeActive && 'is-active')} onClick={() => setPreviewBeforeActive(false)}>
                          <span className="dex-ui-choice-body">
                            <span className="dex-ui-choice-title">{isDe ? '… komplett unsichtbar' : '… completely hidden'} <span className="dex-ui-label-optional">{isDe ? '(Standard)' : '(default)'}</span></span>
                            <span className="dex-ui-choice-desc">{isDe ? 'Nur Admins, Organizer und Test-Team sehen es vorher.' : 'Only admins, organizers and the test team see it beforehand.'}</span>
                          </span>
                          <span className="dex-ui-choice-check">{!previewBeforeActive && <Check size={12} />}</span>
                        </button>
                        <button type="button" role="radio" aria-checked={previewBeforeActive} className={cx('dex-ui-choice', previewBeforeActive && 'is-active')} onClick={() => setPreviewBeforeActive(true)}>
                          <span className="dex-ui-choice-body">
                            <span className="dex-ui-choice-title">{isDe ? '… schon als Vorschau sichtbar' : '… already shown as a preview'}</span>
                            <span className="dex-ui-choice-desc">{isDe ? 'Steht mit dem Hinweis „Anmeldung ab …“ in der Event-Liste; die Anmeldeseite öffnet sich erst ab dem Aktivierungszeitpunkt.' : 'Listed with the note „Registration opens …“; the registration page only opens from the activation time onwards.'}</span>
                          </span>
                          <span className="dex-ui-choice-check">{previewBeforeActive && <Check size={12} />}</span>
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              )}
              </div>
  );
};
