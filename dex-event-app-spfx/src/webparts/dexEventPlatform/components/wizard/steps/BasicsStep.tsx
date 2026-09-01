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
import { Plus, X } from '../../Icons';
import { Icon } from '@fluentui/react/lib/Icon';
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
  currentStep: number;
  currentUser: import("../../../types/index").User;
  dayKeyOfDate: (d: Date) => string;
  description: string;
  DRAFT_KEY: string;
  draftSavedAt: number;
  editEvent: import("../../../types/index").DeloitteEvent;
  emailLogoPreview: any;
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
  setEmailLogoPreview: React.Dispatch<any>;
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
  const { activeFrom, activeScopeIdx, applyDraftPayload, applyEventTemplate, childEventsOf, childTermSingular, currentStep, currentUser, dayKeyOfDate, description, DRAFT_KEY, draftSavedAt, editEvent, emailLogoPreview, errorBorderStyle, events, fieldHasError, fileToBase64, imageBanner, imageDisplay, imageDisplayOpen, imageEditOpen, imageFile, imageOrigFile, imagePreview, imageUploadError, isDe, isEditMode, isFictive, location, logoCropTarget, noDescription, outlookLogoPreview, patchScopeSub, pendingDraft, previewBeforeActive, renderStepIntro, scAllDay, scDescription, scEnd, scImagePreview, scopeSub, scShowAsFree, scStart, scTitle, setActiveFrom, setDescription, setEmailLogoFromPhoto, setEmailLogoPreview, setEventImageUrl, setHtmlEditorMode, setHtmlEditorOpen, setImageBanner, setImageDisplay, setImageDisplayOpen, setImageEditOpen, setImageFile, setImageOrigAspect, setImageOrigFile, setImagePreview, setImageUploadError, setIsFictive, setLogoCropTarget, setNoDescription, setOutlookLogoFromPhoto, setOutlookLogoPreview, setPendingDraft, setPreviewBeforeActive, setScAllDay, setScEnd, setScShowAsFree, setScStart, setScTitle, setShowDemoVariantModal, setShowTemplatePicker, setSubEvents, setSubImageCropIdx, showTemplatePicker, shrinkLogoB64, startDate, subEvents, subEventsOnlyMode, t, templateLoadingId, title, wizardImgAspect, zebraS3Bg } = p;
  return (
              <div style={{ display: visible ? 'block' : 'none' }}>
              {/* v23.6: Demo-Button sitzt jetzt IM grünen Schritt-1-Header
                  (oben rechts), nicht mehr in einer eigenen Zeile darüber. */}
              <h2 className="dex-step-head-title" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12 }}>
                <span>{isDe ? 'Schritt 1 — Grundlagen' : 'Step 1 — Basics'}</span>
                {!isEditMode && (
                  <button
                    type="button"
                    data-tour="wizard-demo"
                    onClick={() => setShowDemoVariantModal(true)}
                    title={isDe ? 'Demo-Vorlage auswählen' : 'Choose demo template'}
                    style={{
                      flexShrink: 0, background: '#fff', color: 'var(--dex-green-dark, #4a7c1f)',
                      border: 'none', borderRadius: 999, padding: '4px 14px',
                      fontSize: '0.8rem', fontWeight: 700, cursor: 'pointer',
                    }}
                  >
                    {isDe ? 'Demo' : 'Demo'}
                  </button>
                )}
              </h2>
              <p className="dex-step-head-lead">
                {isDe
                  ? 'Hier definierst du das Fundament des Events: Titel, Datum, Beschreibung und Bild.'
                  : 'Here you define the foundation of the event: title, date, description and image.'}
              </p>

              {/* v28.89: Alles zwischen hier und dem Titel-Feld gilt für das
                  GESAMTE Event — ob es Sub-Events gibt, wie sie heißen, wie
                  angemeldet wird, Entwurf/Aktivierung. Auf einem Sub-Event-
                  Reiter wäre das falsch am Platz (man würde die Grundsatzfrage
                  „unter" einem einzelnen Termin beantworten), deshalb blenden
                  wir es dort aus und sagen, wo es steht. */}
              {activeScopeIdx > 0 && (
                <WizardHint
                  isDe={isDe}
                  variant="description"
                  title={isDe ? 'Du bearbeitest die Grundlagen eines Sub-Events' : 'You are editing a sub-event’s basics'}
                  style={{ marginBottom: 12 }}
                >
                  {isDe
                    ? <>Titel, Zeiten, Beschreibung und Bild unten gehören zu diesem <strong>{childTermSingular || 'Sub-Event'}</strong>. Die Angaben zum gesamten Event — ob es Sub-Events gibt, wie sie heißen, wie angemeldet wird sowie Entwurf und Aktivierung — stehen auf dem Reiter <strong>{subEventsOnlyMode ? 'Klammer' : 'Haupt-Event'}</strong> oben.</>
                    : <>Title, times, description and image below belong to this <strong>{childTermSingular || 'sub-event'}</strong>. The settings for the event as a whole — whether it has sub-events, how they are named, how people register, plus draft and activation — live on the <strong>{subEventsOnlyMode ? 'bracket' : 'main event'}</strong> tab above.</>}
                </WizardHint>
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
                  <div style={{ margin: '0 0 22px', border: '2px solid var(--dex-green, #86bc25)', borderRadius: 16, background: 'linear-gradient(135deg, rgba(134,188,37,0.10), rgba(0,118,168,0.06))', overflow: 'hidden' }}>
                    {!showTemplatePicker ? (
                      <button
                        type="button"
                        onClick={() => setShowTemplatePicker(true)}
                        style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 20, padding: '18px 22px', background: 'transparent', border: 'none', cursor: 'pointer', textAlign: 'left' }}
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
                          <div style={{ fontWeight: 800, fontSize: '1.05rem', color: 'var(--dex-green-dark, #4a7c1f)', marginBottom: 4 }}>
                            {isDe ? 'Eigenes Event als Vorlage nutzen?' : 'Use one of your events as a template?'}
                          </div>
                          <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                            {isDe
                              ? <>Übernimm Einstellungen und Bild aus einem deiner <strong>{tmpl.length}</strong> bisherigen Events — Datum und Anmeldungen legst du danach neu fest. <span style={{ color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700 }}>Klicken zum Auswählen ▸</span></>
                              : <>Reuse settings and image from one of your <strong>{tmpl.length}</strong> past events. <span style={{ color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700 }}>Click to choose ▸</span></>}
                          </div>
                        </div>
                      </button>
                    ) : (
                      <div style={{ padding: '16px 20px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, marginBottom: 12 }}>
                          <span style={{ fontWeight: 800, fontSize: '1rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
                            {isDe ? 'Vorlage wählen' : 'Choose a template'}
                          </span>
                          <button type="button" onClick={() => setShowTemplatePicker(false)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-500)', fontSize: '0.82rem', fontWeight: 600 }}>
                            {isDe ? 'Abbrechen' : 'Cancel'}
                          </button>
                        </div>
                        <div style={{ display: 'flex', gap: 12, overflowX: 'auto', paddingBottom: 6 }}>
                          {tmpl.map(e => (
                            <button
                              key={e.id}
                              type="button"
                              disabled={!!templateLoadingId}
                              onClick={() => { void applyEventTemplate(e); }}
                              style={{
                                flex: '0 0 auto', width: 150, textAlign: 'left', cursor: templateLoadingId ? 'wait' : 'pointer',
                                background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 12, padding: 0, overflow: 'hidden',
                                boxShadow: '0 2px 8px rgba(0,0,0,0.06)',
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
                return (
                  <div style={{ margin: '0 0 22px', border: '2px solid var(--dex-orange, #ed8b00)', borderRadius: 16, background: 'linear-gradient(135deg, rgba(237,139,0,0.08), rgba(0,118,168,0.05))', padding: '18px 22px', display: 'flex', alignItems: 'center', gap: 20, flexWrap: 'wrap' }}>
                    <div style={{ flexShrink: 0, width: 64, height: 80, borderRadius: 10, border: '3px solid #fff', boxShadow: '0 4px 10px rgba(0,0,0,0.18)', background: 'linear-gradient(135deg, var(--dex-orange, #ed8b00), var(--dex-blue, #0076a8))', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                      <svg width={30} height={30} viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round" xmlns="http://www.w3.org/2000/svg">
                        <path d="M12 20h9" /><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z" />
                      </svg>
                    </div>
                    <div style={{ flex: 1, minWidth: 220 }}>
                      <div style={{ fontWeight: 800, fontSize: '1.05rem', color: '#b86700', marginBottom: 4 }}>
                        {isDe ? 'Aktueller Entwurf' : 'Current draft'}
                      </div>
                      <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                        {isDe
                          ? <><strong>&bdquo;{dTitle}&ldquo;</strong> — zwischengespeichert am {when}. Hochgeladene Bilder sind im Entwurf nicht enthalten und müssten neu gewählt werden.</>
                          : <><strong>&bdquo;{dTitle}&ldquo;</strong> — auto-saved on {when}. Uploaded images are not part of the draft and would need to be re-selected.</>}
                      </div>
                    </div>
                    <div style={{ display: 'flex', gap: 10, flexShrink: 0, flexWrap: 'wrap' }}>
                      <button
                        type="button"
                        className="btn btn-primary"
                        style={{ fontSize: '0.85rem', padding: '8px 18px' }}
                        onClick={() => {
                          try { applyDraftPayload(pendingDraft.data); } catch (err) { console.warn('[DEX] Entwurf-Wiederherstellung fehlgeschlagen:', err); }
                          setPendingDraft(null);
                        }}
                      >
                        {isDe ? 'Entwurf fortsetzen' : 'Continue draft'}
                      </button>
                      <button
                        type="button"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.85rem', padding: '8px 18px' }}
                        onClick={() => {
                          try { localStorage.removeItem(DRAFT_KEY); } catch { /* */ }
                          setPendingDraft(null);
                        }}
                      >
                        {isDe ? 'Entwurf löschen' : 'Delete draft'}
                      </button>
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
                  '1. Als Entwurf speichern — Event nur für Admins, Organizer und Test-Team sichtbar; optional Aktiv-Ab-Datum für automatisches Go-Live',
                  '2. Event-Titel',
                  '3. Datum (Start &amp; Ende) — füllt die Anmelde- und Storno-Deadlines automatisch vor',
                  '4. Beschreibung (optional, HTML-Editor)',
                  '5. Event-Bild hochladen — oben auf der Detailseite und in den Mails verwendet',
                  '6. Organizer auswählen — bekommen alle Organizer-Mails',
                  '7. Test-Team — sieht das Event schon im Entwurfsmodus',
                  '8. Check-In Team — darf nur das QR-/Check-In-Tool nutzen',
                ],
                [
                  '1. Save as draft — visible only to admins, organizers, and the test team; optional active-from date for automatic go-live',
                  '2. Event title',
                  '3. Date (start &amp; end) — pre-fills the registration and cancellation deadlines',
                  '4. Description (optional, HTML editor)',
                  '5. Upload an event image — shown at the top of the detail page and in emails',
                  '6. Pick organizers — they receive all organizer emails',
                  '7. Test team — can see the event already in draft mode',
                  '8. Check-in team — may only use the QR / check-in tool',
                ]
              )}

              {/* v9.21: Entwurf-Flag als erster Schritt — vor Title.
                  Default ist on, der Organizer kann die Test-Strecke
                  in Ruhe aufbauen, das Test-Team durchspielen lassen,
                  und ohne Aengste sein Event entwickeln.
                  v22.27: volle Breite wie die übrigen Hinweis-Boxen
                  (vorher maxWidth 720). */}
              <div className="form-group" style={{ marginTop: 0, marginBottom: 20, paddingBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', padding: 14, background: isFictive ? 'rgba(237,139,0,0.06)' : 'var(--dex-gray-50, #f8f9fa)', borderRadius: 'var(--dex-radius, 12px)', border: `1px solid ${isFictive ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-gray-200)'}` }}>
                  <StepBadge n={1} />
                  <input
                    type="checkbox"
                    checked={isFictive}
                    onChange={e => setIsFictive(e.target.checked)}
                    style={{ width: 18, height: 18, cursor: 'pointer', marginTop: 3 }}
                  />
                  <span style={{ fontSize: '0.9rem' }}>
                    <strong>{t('create.fictive')}</strong>
                    <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-500)', lineHeight: 1.5, marginTop: 4 }}>
                      {t('create.fictive.hint')}
                    </span>
                  </span>
                </label>
                {/* v9.21: ActiveFrom direkt unter dem Entwurfs-Toggle — wenn
                    der Organizer ein Live-Datum setzt, geht das Event ab dann
                    auch wenn das Entwurf-Häkchen noch on ist. Optional. */}
                <div style={{ marginTop: 12, paddingLeft: 4 }}>
                  <label style={{ display: 'block', fontSize: '0.85rem', color: 'var(--dex-gray-700)', marginBottom: 4, fontWeight: 500 }}>
                    Aktiv ab (optional)
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
                    timeCaption="Uhrzeit"
                    dateFormat="dd.MM.yyyy, HH:mm"
                    locale="de"
                    placeholderText="Datum und Uhrzeit wählen"
                    className="form-input"
                    wrapperClassName="dex-datepicker-wrapper"
                    calendarClassName="dex-datepicker-calendar"
                    popperPlacement="bottom-start"
                    isClearable
                    autoComplete="off"
                  />
                  {/* v23.14: Vorschau-Wahl — nur sinnvoll bei gesetztem „Aktiv ab". */}
                  {activeFrom && (
                    <div style={{ marginTop: 12, padding: '12px 14px', background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                      <div style={{ fontSize: '0.85rem', fontWeight: 600, color: 'var(--dex-gray-800)', marginBottom: 6 }}>
                        {isDe ? 'Vor dem Aktivierungszeitpunkt …' : 'Before the activation time …'}
                      </div>
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', fontSize: '0.84rem', marginBottom: 6 }}>
                        <input type="radio" name="previewBeforeActive" checked={!previewBeforeActive} onChange={() => setPreviewBeforeActive(false)} style={{ marginTop: 3 }} />
                        <span>{isDe
                          ? <>… <strong>komplett unsichtbar</strong> für Teilnehmer (Standard). Nur Admins/Organizer/Test-Team sehen es vorher.</>
                          : <>… <strong>completely hidden</strong> from attendees (default). Only admins/organizers/test team see it beforehand.</>}</span>
                      </label>
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', fontSize: '0.84rem' }}>
                        <input type="radio" name="previewBeforeActive" checked={previewBeforeActive} onChange={() => setPreviewBeforeActive(true)} style={{ marginTop: 3 }} />
                        <span>{isDe
                          ? <>… schon als <strong>Vorschau sichtbar</strong> in der Event-Liste — mit dem Hinweis „Anmeldung ab …“. Die Anmeldeseite lässt sich aber erst ab dem Aktivierungszeitpunkt öffnen.</>
                          : <>… already shown as a <strong>preview</strong> in the event list — with the note „Registration opens …“. The registration page can only be opened from the activation time onwards.</>}</span>
                      </label>
                    </div>
                  )}
                </div>
              </div>

              </>)}{/* v28.89: Ende der event-weiten Angaben */}

              <div className="form-group" style={{ paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={2} />
                  {!scopeSub && <span className="required">*</span>} {t('create.eventtitle')}
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
                  placeholder={scopeSub ? t('create.subevents.title.placeholder') : 'z.B. Sommerfest 2026'}
                  style={scopeSub ? undefined : errorBorderStyle('title')}
                />
                {!scopeSub && fieldHasError('title') && <span style={{ color: 'var(--dex-red)', fontSize: '0.75rem' }}>{t('create.error.required')}</span>}
              </div>

              {/* v9.24: Event-Datum direkt nach Title — auto-fillt die Deadlines.
                  Vorher in Step 1, jetzt in Step 0 weil das fundamentale Info ist. */}
              <div className="form-group" style={{ paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={3} />
                  Datum (Start &amp; Ende)
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
                </label>
              <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                <div className="form-group" style={{ marginBottom: 0 }}>
                  <label className="form-label">
                    {!scopeSub && <span className="required">*</span>} {t('create.startdate')}
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
                    timeCaption="Uhrzeit"
                    dateFormat={scAllDay ? 'dd.MM.yyyy' : 'dd.MM.yyyy, HH:mm'}
                    locale="de"
                    // v28.66: Beim Sub-Event heißt leer „Zeit des Hauptevents".
                    placeholderText={scopeSub ? t('create.subevents.time.placeholder') : 'Datum und Uhrzeit wählen'}
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
                <div className="form-group" style={{ marginBottom: 0 }}>
                  <label className="form-label">
                    {!scopeSub && <span className="required">*</span>} {t('create.enddate')}
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
                    timeCaption="Uhrzeit"
                    dateFormat={scAllDay ? 'dd.MM.yyyy' : 'dd.MM.yyyy, HH:mm'}
                    locale="de"
                    placeholderText={scopeSub ? t('create.subevents.time.placeholder') : 'Datum und Uhrzeit wählen'}
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
              <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', marginTop: 10 }}>
                <input
                  type="checkbox"
                  checked={scAllDay}
                  onChange={e => setScAllDay(e.target.checked)}
                  style={{ width: 18, height: 18, marginTop: 1, flexShrink: 0, cursor: 'pointer' }}
                />
                <span style={{ fontSize: '0.9rem' }}>
                  <strong>{isDe ? 'Ganztägiger Termin' : 'All-day event'}</strong>
                  <span style={{ display: 'block', color: 'var(--dex-gray-600)', marginTop: 2, fontWeight: 400 }}>
                    {isDe
                      ? 'Der Outlook-Termin erscheint dann oben im Kalenderkopf statt als Block über den Tag — und lässt die Verfügbarkeit der Teilnehmer frei. Ohne Haken bucht ein Termin von 00:00 bis 23:59 den kompletten Tag als belegt.'
                      : 'The Outlook entry then appears in the calendar header instead of as a block across the day — and leaves attendees shown as free. Without it, a 00:00–23:59 entry books the whole day as busy.'}
                  </span>
                </span>
              </label>
              {/* v29.54: Kalender blockieren ja/nein. Der Haken ist ANGEHAKT =
                  beschäftigt (Default, bisheriges Verhalten); gespeichert wird
                  der umgekehrte Wert `showAsFree` — siehe Kommentar an
                  DeloitteEvent.showAsFree. Die Umkehrung passiert genau hier,
                  an einer Stelle, und nirgends sonst. */}
              <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', marginTop: 10 }}>
                <input
                  type="checkbox"
                  checked={!scShowAsFree}
                  onChange={e => setScShowAsFree(!e.target.checked)}
                  style={{ width: 18, height: 18, marginTop: 1, flexShrink: 0, cursor: 'pointer' }}
                />
                <span style={{ fontSize: '0.9rem' }}>
                  <strong>{isDe ? 'Termin blockiert den Kalender' : 'Entry blocks the calendar'}</strong>
                  <span style={{ display: 'block', color: 'var(--dex-gray-600)', marginTop: 2, fontWeight: 400 }}>
                    {isDe
                      ? <>Der Termin steht bei den Teilnehmern auf <strong>Beschäftigt</strong> — Kollegen sehen sie als nicht verfügbar. {scAllDay ? <>Bei einem ganztägigen Termin gilt damit der <strong>komplette Arbeitstag</strong> als belegt; für ein Angebot, zu dem man nur zeitweise dazukommt, nimmst du den Haken besser raus.</> : <>Ohne Haken erscheint der Termin als <strong>Frei</strong> und blockiert nichts.</>}</>
                      : <>The entry shows as <strong>busy</strong> for attendees — colleagues see them as unavailable. {scAllDay ? <>For an all-day entry that marks the <strong>entire working day</strong> as taken; for something people only drop into, better untick it.</> : <>Without it the entry shows as <strong>free</strong> and blocks nothing.</>}</>}
                  </span>
                </span>
              </label>
              <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginTop: 8, marginBottom: 0 }}>
                {scAllDay
                  ? (isDe
                    ? 'Ganztägig: Es zählt nur das Datum — die Uhrzeit spielt für den Outlook-Termin keine Rolle mehr.'
                    : 'All-day: only the date matters — the time is no longer used for the Outlook entry.')
                  : scopeSub
                    ? (isDe
                      ? 'Leer lassen heißt: Dieses Sub-Event übernimmt die Zeiten des Hauptevents. Die Uhrzeit wird für den Outlook-Kalendereintrag der Teilnehmer verwendet.'
                      : 'Leaving these empty means the sub-event inherits the main event’s times. The time is used for the attendees’ Outlook entry.')
                    : 'Die Uhrzeit wird für den Outlook-Kalendereintrag der Teilnehmer verwendet.'}
              </p>
              </div>

              <div className="form-group" style={{ paddingBottom: 20, marginBottom: 20, borderBottom: '1px solid var(--dex-gray-100)' }}>
                {/* v28.7: kein <label> mehr, sondern <div> — rechts sitzt jetzt
                    die „Keine Beschreibung"-Checkbox mit eigenem <label>;
                    verschachtelte Labels würden Klicks auf die Überschrift
                    fälschlich auf die Checkbox umleiten. */}
                <div className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                  <StepBadge n={4} />
                  {t('create.description')}
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
                  {!scopeSub && (
                    <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, marginLeft: 'auto', fontSize: '0.78rem', fontWeight: 400, color: 'var(--dex-gray-600)', cursor: 'pointer', whiteSpace: 'nowrap' }}>
                      <input
                        type="checkbox"
                        checked={noDescription}
                        onChange={e => {
                          const on = e.target.checked;
                          setNoDescription(on);
                          if (on) setDescription('');
                        }}
                        style={{ accentColor: 'var(--dex-green)' }}
                      />
                      {isDe ? 'Keine Beschreibung nutzen' : 'Don’t use a description'}
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
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', margin: 0 }}>
                    {isDe ? 'Auf der Anmelde-Seite wird keine Beschreibung angezeigt.' : 'No description will be shown on the registration page.'}
                  </p>
                ) : (
                <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
                  <button
                    type="button"
                    className="btn btn-secondary"
                    onClick={() => { setHtmlEditorMode('description'); setHtmlEditorOpen(true); }}
                    style={{ fontSize: '0.85rem' }}
                  >
                    {isDe ? 'Bearbeiten & Vorschau' : 'Edit & Preview'}
                  </button>
                  <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', flex: 1, minWidth: 200 }}>
                    {scDescription
                      ? `${scDescription.replace(/<[^>]+>/g, '').substring(0, 120)}${scDescription.length > 120 ? '…' : ''}`
                      : (isDe ? 'Keine Beschreibung gesetzt — klicke „Bearbeiten" zum Hinzufügen.' : 'No description set — click „Edit" to add one.')}
                  </span>
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

              <div className="form-group">
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={5} />
                  {t('create.eventimage')}
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
                </label>
                {scImagePreview && (
                  <div style={{ position: 'relative', marginBottom: 8, display: 'block', width: 'fit-content', maxWidth: '100%' }}>
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
                    <button
                      type="button"
                      onClick={() => {
                        // v28.89: Auf einem Sub-Event-Reiter betrifft das
                        // Entfernen dessen Bild — imageRemoved ist das Signal
                        // für den Save, das gespeicherte Attachment zu löschen.
                        if (scopeSub) { patchScopeSub({ imagePreview: '', imageFile: null, imageRemoved: true }); return; }
                        setImageFile(null); setImagePreview(''); setEventImageUrl(''); setImageOrigFile(null); setImageOrigAspect(null);
                      }}
                      style={{
                        position: 'absolute', top: 8, right: 8, background: 'rgba(0,0,0,0.6)',
                        color: '#fff', border: 'none', borderRadius: '50%', width: 28, height: 28,
                        cursor: 'pointer', fontSize: '1rem', display: 'flex', alignItems: 'center', justifyContent: 'center',
                      }}
                    >
                      <X size={14} />
                    </button>
                    {/* v23.15: Bild editieren (zuschneiden / auf Kreis). */}
                    <button
                      type="button"
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
                      style={{
                        position: 'absolute', bottom: 8, right: 8, background: 'rgba(0,0,0,0.6)',
                        color: '#fff', border: 'none', borderRadius: 999, padding: '4px 12px',
                        cursor: 'pointer', fontSize: '0.78rem', fontWeight: 600,
                        display: 'inline-flex', alignItems: 'center', gap: 6,
                      }}
                    >
                      <Icon iconName="Crop" style={{ fontSize: 13 }} /> {isDe ? 'Bild editieren' : 'Edit image'}
                    </button>
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
                  <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 8, overflow: 'hidden' }}>
                    <button
                      type="button"
                      onClick={() => setImageDisplayOpen(o => !o)}
                      style={{ width: '100%', textAlign: 'left', background: 'var(--dex-gray-50, #f7f7f5)', border: 'none', cursor: 'pointer', padding: '10px 12px', fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700)', display: 'flex', alignItems: 'center', gap: 8 }}
                    >
                      <Icon iconName={imageDisplayOpen ? 'ChevronDown' : 'ChevronRight'} style={{ fontSize: 12 }} />
                      {isDe ? 'Darstellung pro Ansicht (optional)' : 'Per-view display (optional)'}
                    </button>
                    {imageDisplayOpen && (
                      <div style={{ padding: '12px 14px' }}>
                        <p style={{ margin: '0 0 12px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                          {isDe
                            ? 'Optional: Du kannst das Bild pro Ansicht unterschiedlich zoomen und vertikal verschieben, damit es überall gut sitzt. Standard (nichts einstellen) = das Bild füllt den Bereich zentriert.'
                            : 'Optional: zoom and vertically position the image differently per view so it sits well everywhere. Default (leave untouched) = the image fills the area centered.'}
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
                                  <button type="button" className="btn btn-secondary" style={{ fontSize: '0.74rem', padding: '3px 10px', marginTop: 4 }} onClick={() => setImageDisplay(prev => { const n = { ...prev }; delete n[view.key]; return n; })}>
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
                                <button type="button" className="btn btn-secondary" style={{ fontSize: '0.74rem', padding: '3px 10px', marginTop: 4 }} onClick={() => setImageDisplay(prev => { const n = { ...prev }; delete n[view.key]; return n; })}>
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
                <label style={{
                  display: 'inline-flex', alignItems: 'center', gap: 8,
                  padding: '8px 16px', borderRadius: 'var(--dex-radius)',
                  border: '2px dashed var(--dex-gray-300)', cursor: 'pointer',
                  fontSize: '0.85rem', color: 'var(--dex-gray-600)',
                  transition: 'border-color 0.2s, background 0.2s',
                }}>
                  <Plus size={16} />
                  {(scopeSub ? scopeSub.imageFile : imageFile)?.name || (isDe ? 'Bild auswählen' : 'Choose image')}
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
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginTop: 6, marginBottom: 0 }}>
                    {isDe
                      ? 'Erscheint als Vorschaubild neben diesem Sub-Event in der Auswahl auf der Anmeldeseite. Ohne eigenes Bild bleibt die Zeile dort schlicht ohne Vorschau.'
                      : 'Shown as a thumbnail next to this sub-event in the selection on the registration page. Without its own image the row simply has no preview.'}
                  </p>
                )}
                {imageUploadError && (
                  <p style={{ color: 'var(--dex-red, #c00)', fontSize: '0.8rem', marginTop: 4 }}>{imageUploadError}</p>
                )}
                {/* v28.90: Pflichtanmeldung — eine Einstellung DIESES
                    Sub-Events, deshalb hier bei seinen Grundlagen und nicht
                    (neunmal wiederholt) in der Liste auf der Klammer-Ebene.
                    v28.77: Der Haken wurde als „dieses Sub-Event ist buchbar"
                    missverstanden und darum bei ALLEN gesetzt — das Ergebnis
                    ist das Gegenteil einer Auswahl. Diesen Zustand benennen,
                    sobald er eintritt, mit einem Klick zum Zurücknehmen. */}
                {scopeSub && (
                  <div style={{ marginTop: 20, paddingTop: 20, borderTop: '1px solid var(--dex-gray-100)' }}>
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, padding: '10px 12px', borderRadius: 8, border: `1px solid ${scopeSub.mandatory ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`, background: scopeSub.mandatory ? 'rgba(134,188,37,0.06)' : '#fff', cursor: 'pointer' }}>
                      <input
                        type="checkbox"
                        checked={!!scopeSub.mandatory}
                        onChange={e => patchScopeSub({ mandatory: e.target.checked })}
                        style={{ width: 18, height: 18, marginTop: 1, flexShrink: 0, cursor: 'pointer' }}
                      />
                      <span style={{ fontSize: '0.85rem' }}>
                        <strong>{isDe ? `Pflichtanmeldung für dieses ${childTermSingular || 'Sub-Event'}` : `Mandatory registration for this ${childTermSingular || 'sub-event'}`}</strong>
                        <span style={{ display: 'block', color: 'var(--dex-gray-600)', marginTop: 2, fontWeight: 400 }}>
                          {isDe
                            ? <>Wenn aktiv, muss jeder Teilnehmer diesen Termin mitbuchen — eine Anmeldung ohne ihn ist dann nicht möglich. <strong>Nur setzen, wenn er für wirklich alle verpflichtend ist</strong> (z.B. eine Auftaktveranstaltung). Für &bdquo;darf gebucht werden&ldquo; ist der Haken nicht nötig.</>
                            : <>If active, every attendee must include this date — registering without it is then not possible. <strong>Only set this if it is truly compulsory for everyone</strong> (e.g. a kick-off). For &bdquo;may be booked&ldquo; the checkbox is not needed.</>}
                        </span>
                      </span>
                    </label>
                    {(() => {
                      const named = subEvents.filter(s => (s.title || '').trim());
                      const mandatoryCount = named.filter(s => s.mandatory).length;
                      if (named.length < 2 || mandatoryCount !== named.length || !scopeSub.mandatory) return null;
                      return (
                        <div style={{
                          margin: '10px 0 0', padding: '9px 11px', borderRadius: 8,
                          background: '#fff8e6', border: '1px solid #e0b34d', color: '#7a5a12',
                          fontSize: '0.78rem', lineHeight: 1.55,
                        }}>
                          <strong>{isDe ? `Alle ${named.length} Sub-Events sind als Pflicht markiert` : `All ${named.length} sub-events are marked mandatory`}</strong>
                          <div style={{ marginTop: 3 }}>
                            {isDe
                              ? <>Damit gibt es faktisch <strong>keine Auswahl mehr</strong> — wer teilnehmen möchte, muss <strong>alle {named.length}</strong> mitbuchen; wer auch nur einen Termin nicht kann, kann sich gar nicht anmelden.{subEventsOnlyMode ? <> Bei diesem Klammerevent läuft die Anmeldung ohnehin ausschließlich über die Sub-Events — der Haken ist dafür <strong>nicht nötig</strong>.</> : null} Gemeint war vermutlich, dass die Sub-Events buchbar sind — dafür lässt du den Haken einfach weg.</>
                              : <>That leaves <strong>no choice at all</strong> — attendees must book <strong>all {named.length}</strong>; anyone unavailable for a single date cannot register.{subEventsOnlyMode ? <> For this bracket event registration runs via the sub-events anyway — the checkbox is <strong>not needed</strong> for that.</> : null}</>}
                          </div>
                          <button
                            type="button"
                            className="btn btn-primary"
                            style={{ fontSize: '0.78rem', padding: '5px 12px', marginTop: 8 }}
                            onClick={() => setSubEvents(prev => prev.map(s => ({ ...s, mandatory: false })))}
                          >
                            {isDe ? 'Pflicht bei allen entfernen' : 'Remove mandatory from all'}
                          </button>
                        </div>
                      );
                    })()}
                  </div>
                )}
                {/* v28.5: Layout-Wahl fürs Event-Bild auf der Anmeldeseite —
                    Banner in voller Breite ÜBER den Infos (gut für breite
                    Querformat-Fotos) vs. kompakt links neben den Infos.
                    v28.10: nur noch bei Querformat-Bildern (Ratio >= 1.2)
                    anbieten — für Kreis-/Quadrat-/Hochkant-Bilder ergibt
                    das Banner-Layout keinen Sinn. */}
                {!scopeSub && (imagePreview || imageFile) && wizardImgAspect != null && wizardImgAspect >= 1.2 && (
                  <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginTop: 12, padding: '10px 12px', borderRadius: 8, border: `1px solid ${imageBanner ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`, background: imageBanner ? 'rgba(134,188,37,0.06)' : '#fff', cursor: 'pointer' }}>
                    <input
                      type="checkbox"
                      checked={imageBanner}
                      onChange={e => setImageBanner(e.target.checked)}
                      style={{ width: 18, height: 18, marginTop: 1, flexShrink: 0, cursor: 'pointer' }}
                    />
                    <span style={{ fontSize: '0.82rem' }}>
                      <strong>{isDe ? 'Bild als Banner über den Event-Infos anzeigen' : 'Show image as a banner above the event info'}</strong>
                      <span style={{ display: 'block', color: 'var(--dex-gray-600)', marginTop: 2, fontWeight: 400 }}>
                        {isDe
                          ? 'Empfohlen für breite Querformat-Fotos: Das Bild liegt auf der Anmeldeseite in voller Kartenbreite oben, Titel/Datum/Ort folgen darunter. Aus = Bild sitzt kompakt links neben den Infos (Standard).'
                          : 'Recommended for wide landscape photos: the image spans the full card width at the top of the registration page, with title/date/location below. Off = compact image to the left of the info (default).'}
                      </span>
                    </span>
                  </label>
                )}
              </div>


              </div>
  );
};
