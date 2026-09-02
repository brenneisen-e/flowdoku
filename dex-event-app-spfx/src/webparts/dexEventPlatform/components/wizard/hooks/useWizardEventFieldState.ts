/* useWizardEventFieldState — aus EventCreationPage.tsx ausgelagert (Zeilen 568-1023 des
 * damaligen Stands). Der Bereich ist zeichengleich uebernommen, inklusive der
 * Reihenfolge seiner Hook-Aufrufe: der Aufruf dieses Hooks steht in der Seite
 * an genau der Stelle, an der der Bereich vorher stand. Schwerpunkt: Felder des Events (Ort, Zeiten, Beschreibung, Bild).
 *
 * Die Grenze ist mechanisch gezogen (zusammenhaengender Bereich), nicht
 * thematisch — der Name beschreibt den Schwerpunkt, nicht eine reine Trennung. */
import * as React from 'react';
import { buildOutlookLocation } from '../../../utils/eventFormat';
import { EventType } from '../../../types';
import { ImgView, SubEventDraft } from '../../wizard/wizardTypes';
import { BundledComm, bundledCommOf } from '../../../utils/bundledComm';
import { CustomFieldInput } from '../../wizard/customFieldInput';
import { reinsertOrganizerPlaceholder } from '../../wizard/wizardHelpers';
import { parseOutlookHeadings, stripOutlookWrapper } from '../../../services/EmailTemplates';

export interface UseWizardEventFieldStateCtx {
  editEvent: import("../../../types/index").DeloitteEvent;
  ensureEventDocuments: (eventIds: string[]) => Promise<void>;
  isoToLocal: (iso: string) => string;
  locale: import("../../../context/LanguageContext").Locale;
}

export function useWizardEventFieldState(ctx: UseWizardEventFieldStateCtx) {
  const { editEvent, ensureEventDocuments, isoToLocal, locale } = ctx;
  const [outlookLocationOverride, setOutlookLocationOverride] = React.useState<string>(() => {
    if (!editEvent) return '';
    const auto = buildOutlookLocation(editEvent.location, editEvent.locationAddress);
    const stored = editEvent.outlookLocation || '';
    return (stored && stored !== auto) ? stored : '';
  });
  const [locationFilter, setLocationFilter] = React.useState(
    editEvent ? editEvent.locationAudience.join(', ') : ''
  );
  const [audience, setAudience] = React.useState(
    editEvent && editEvent.audienceFilter ? editEvent.audienceFilter.join(', ') : ''
  );
  // Default für neue Events: 'OR' — konsistent mit EventContext-Read-Fallback
  // und konservativer (UND-Verknüpfung kann Mitarbeiter unbeabsichtigt
  // ausschliessen). Bestehende Events behalten ihren gespeicherten Wert.
  const [filterMode, setFilterMode] = React.useState<'AND' | 'OR'>(
    editEvent ? editEvent.filterMode : 'OR'
  );
  const [description, setDescription] = React.useState(editEvent ? editEvent.description : '');
  // v28.7: „Keine Beschreibung nutzen" — reiner UI-Schalter im Wizard
  // (Default: Beschreibung nutzen). Anhaken leert die Beschreibung und
  // blendet den Editor-Zugang aus; gespeichert wird schlicht ''.
  const [noDescription, setNoDescription] = React.useState<boolean>(() => {
    // v28.79: beim Bearbeiten aus dem gespeicherten Flag vorbelegen —
    // sonst stand der Schalter nach dem Neuladen wieder auf „Beschreibung
    // nutzen", obwohl der Organizer sie bewusst weggelassen hatte.
    if (!editEvent) return false;
    if ((editEvent.description || '').trim()) return false;
    try {
      const ov = JSON.parse(editEvent.emailTemplateOverrides || '{}');
      return !!(ov && ov._noDescription);
    } catch { return false; }
  });
  // EventType wird nicht mehr als UI-Feld abgefragt (v5.2) — neue Events:
  // aus Template abgeleitet (b2run → 'B2Run', sonst → 'Other'). Bei Edit:
  // den gespeicherten Wert beibehalten. Die Variable wird weiterhin für
  // Card-Gradient + B2Run-spezifische Admin-Funktionen gebraucht.
  const [storedEventType] = React.useState<EventType>(editEvent ? editEvent.type : 'Other');
  const [startDate, setStartDate] = React.useState(editEvent ? isoToLocal(editEvent.startDate) : '');
  const [endDate, setEndDate] = React.useState(editEvent ? isoToLocal(editEvent.endDate) : '');
  const [registrationDeadline, setRegistrationDeadline] = React.useState(
    editEvent ? isoToLocal(editEvent.registrationDeadline) : ''
  );
  // v28.20: EXPLIZITE Anmeldefrist der Klammer (Piggyback _klammerDeadline).
  // Optional — leer heißt wie bisher: offen, solange ein Sub-Event offen ist.
  // Gesetzt + abgelaufen = Anmeldung fürs GESAMTE Event geschlossen.
  const [klammerDeadline, setKlammerDeadline] = React.useState(
    editEvent && editEvent.klammerDeadline ? isoToLocal(editEvent.klammerDeadline) : ''
  );
  const [lastDeregisterDate, setLastDeregisterDate] = React.useState(editEvent ? isoToLocal(editEvent.lastDeregisterDate) : '');
  // v29.25: Selbst-Abmeldung, zweistufig. Stufe 1: „Abmeldung durch User
  // ermöglichen" (Default ja; bei Nein gibt es keine Abmeldefrist und nur
  // Organizer/Admins melden ab — Piggyback _noSelfCancel). Stufe 2 (nur bei
  // Ja mit gesetzter Frist): „auch nach der Abmeldefrist erlauben" (Default
  // ja = Late-Cancel mit Organizer-Mail; bei Nein Piggyback
  // _noCancelAfterDeadline).
  // v29.38: Optionaler Teams-Besprechungslink. DEX legt KEIN Teams-Meeting an —
  // der Organizer fuegt den Link seiner eigenen Besprechung ein. Er landet als
  // Teilnahme-Block im Outlook-Termin.
  const [teamsLink, setTeamsLink] = React.useState<string>(editEvent?.teamsLink || '');
  /**
   * v30.26: Online-Meeting-Modus des Events — die Entscheidung trifft der
   * Organizer pro Event, NICHT der Flow global.
   *
   *  'none' — kein Online-Meeting (Präsenz).
   *  'own'  — der Organizer legt die Besprechung selbst in Outlook/Teams an
   *           und trägt den Link ein (v29.38-Feld). Er behält damit alle
   *           Besprechungsoptionen (Lobby, Aufzeichnung, Referenten).
   *  'auto' — DEX lässt den Flow eine echte Teams-Besprechung erzeugen
   *           (Spalte OutlookIsOnlineMeeting). Bequem und mit „Teilnehmen"-
   *           Knopf im Kalender, ABER der Termin gehört dem Gruppen-/
   *           No-Reply-Postfach: An den Besprechungsoptionen kann danach
   *           niemand mehr etwas ändern.
   *
   * Abgeleitet aus dem gespeicherten Stand, damit Bestandsevents (nur
   * teamsLink gepflegt) unverändert als 'own' erscheinen.
   */
  const [onlineMeetingMode, setOnlineMeetingMode] = React.useState<'none' | 'own' | 'auto'>(() => {
    if (editEvent?.outlookIsOnlineMeeting) return 'auto';
    if ((editEvent?.teamsLink || '').trim()) return 'own';
    return 'none';
  });
  const [userCancelAllowed, setUserCancelAllowed] = React.useState<boolean>(!(editEvent && editEvent.noSelfCancel));
  const [noCancelAfterDeadline, setNoCancelAfterDeadline] = React.useState<boolean>(!!(editEvent && editEvent.noCancelAfterDeadline));
  // v9.22: Auto-Fill der Deadlines wenn Start-Datum gesetzt wird und die
  // Deadlines noch leer sind. Default-Logik:
  //   - RegistrationDeadline: 7 Tage vor Event-Start
  //   - LastDeregisterDate: 3 Tage vor Event-Start
  // Der Organizer kann beides überschreiben — wir aktualisieren NICHT,
  // wenn der User schon einen Wert gesetzt hat.
  const autoFillRanRef = React.useRef(false);
  React.useEffect(() => {
    if (autoFillRanRef.current) return;
    if (!startDate) return;
    if (registrationDeadline || lastDeregisterDate) return;
    try {
      const start = new Date(startDate);
      if (isNaN(start.getTime())) return;
      const fmt = (d: Date): string => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}T${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
      const reg = new Date(start.getTime() - 7 * 24 * 60 * 60 * 1000);
      const lastCancel = new Date(start.getTime() - 3 * 24 * 60 * 60 * 1000);
      setRegistrationDeadline(fmt(reg));
      setLastDeregisterDate(fmt(lastCancel));
      autoFillRanRef.current = true;
    } catch { /* */ }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [startDate]);
  // v26.51: Wird das START-Datum GEÄNDERT (v.a. beim Bearbeiten: Event wird
  // verschoben), wandern gesetzte An-/Abmeldefristen relativ mit — gleiche
  // Logik wie vorher: War die Anmeldefrist 1 Woche vor dem Event, liegt sie
  // nach der Verschiebung wieder 1 Woche vor dem (neuen) Event-Beginn.
  const prevStartForShiftRef = React.useRef(startDate);
  React.useEffect(() => {
    const prev = prevStartForShiftRef.current;
    prevStartForShiftRef.current = startDate;
    if (!prev || !startDate || prev === startDate) return;
    const oldTs = new Date(prev).getTime();
    const newTs = new Date(startDate).getTime();
    if (!isFinite(oldTs) || !isFinite(newTs)) return;
    const delta = newTs - oldTs;
    if (!delta) return;
    const fmt = (d: Date): string => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}T${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
    const shift = (val: string): string => {
      if (!val) return val;
      const t = new Date(val).getTime();
      if (!isFinite(t)) return val;
      return fmt(new Date(t + delta));
    };
    setRegistrationDeadline(v => shift(v));
    setLastDeregisterDate(v => shift(v));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [startDate]);
  const [maxParticipants, setMaxParticipants] = React.useState(
    editEvent && editEvent.maxParticipants ? editEvent.maxParticipants.toString() : ''
  );
  const [unlimitedParticipants, setUnlimitedParticipants] = React.useState(
    !editEvent || !editEvent.maxParticipants || editEvent.maxParticipants === 0
  );
  // v22.37: Neues Event startet UNBEGRENZT → standardmäßig KEINE Warteliste.
  // Erst wenn der Organizer die Teilnehmerzahl begrenzt, wird die Warteliste
  // automatisch aktiviert (Default ja bei begrenzter Kapazität — siehe
  // Unbegrenzt-Toggle-onChange).
  const [waitlistEnabled, setWaitlistEnabled] = React.useState(
    editEvent && typeof editEvent.waitlistEnabled !== 'undefined' ? editEvent.waitlistEnabled : false
  );
  const [eventImageUrl, setEventImageUrl] = React.useState(editEvent ? (editEvent.imageUrl || '') : '');
  const [imageFile, setImageFile] = React.useState<File | null>(null);
  const [imagePreview, setImagePreview] = React.useState(editEvent ? (editEvent.imageUrl || '') : '');
  // v28.11: Frisch hochgeladenes ORIGINAL (vor dem Zuschnitt) + dessen
  // Seitenverhältnis. Wird nur persistiert (zweites Attachment
  // __eventimgorig__ + Piggyback _imageOrigUrl), wenn ein Querformat-
  // Original per Zuschnitt rund/quadratisch wurde — die Event-Liste zeigt
  // dann das Original als Kachel-Hintergrund.
  const [imageOrigFile, setImageOrigFile] = React.useState<File | null>(null);
  const [imageOrigAspect, setImageOrigAspect] = React.useState<number | null>(null);
  // v23.15: Bild-Editor (Zuschneiden / Kreis) offen?
  const [imageEditOpen, setImageEditOpen] = React.useState(false);
  // v26.97: Zuschneiden des Mail-/Outlook-Kopfbildes (nutzt dasselbe
  // ImageCropModal wie das Event-Bild). Ziel = welches Logo gerade zugeschnitten
  // wird ('email' oder 'outlook').
  const [logoCropTarget, setLogoCropTarget] = React.useState<'email' | 'outlook' | null>(null);
  // v28.30: Merker, ob das aktuell gesetzte Kopfbild per „Event-Foto
  // übernehmen" entstanden ist. Nur für die Optik des Knopfs (gruen + Haken =
  // „ist übernommen"). Bewusst NICHT persistiert: nach dem Neuladen zeigt die
  // Vorschau darunter ohnehin das echte Bild, und ein zweiter Klick auf den
  // Knopf ist folgenlos (er setzt dasselbe Bild noch einmal).
  const [emailLogoFromPhoto, setEmailLogoFromPhoto] = React.useState(false);
  const [outlookLogoFromPhoto, setOutlookLogoFromPhoto] = React.useState(false);
  // v27.11: Zuschneiden eines Sub-Event-Bildes — Index des Sub-Events in
  // `subEvents`, dessen Bild gerade im ImageCropModal offen ist (null = zu).
  const [subImageCropIdx, setSubImageCropIdx] = React.useState<number | null>(null);
  // v23.19: Optionale Pro-Ansicht-Darstellung (Zoom + vertikale Position).
  // Default leer = Standard (cover/zentriert) — nur auf Wunsch eingestellt.
  const [imageDisplay, setImageDisplay] = React.useState<{ card?: ImgView; hero?: ImgView }>(editEvent && editEvent.imageDisplay ? editEvent.imageDisplay : {});
  const [imageDisplayOpen, setImageDisplayOpen] = React.useState(false);

  // v29.47: Der Boot lädt die Anhänge nicht mehr mit. Beim Bearbeiten braucht
  // der Wizard sie (Schritt „Dokumente") — hier für das bearbeitete Event
  // nachholen, damit die Liste nicht fälschlich leer aussieht und ein Save die
  // bestehenden Dateien nicht als „entfernt" behandelt.
  React.useEffect(() => {
    if (editEvent?.id) void ensureEventDocuments([editEvent.id]);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editEvent?.id]);

  // v28.5: Bild als Banner über den Event-Infos (statt kompakt links) —
  // Organizer-Wahl, sinnvoll für breite Querformat-Fotos. Piggyback _imageBanner.
  const [imageBanner, setImageBanner] = React.useState<boolean>(!!(editEvent && editEvent.imageBanner));
  // v28.91: Sub-Events sind Termine (ein Tag je Sub-Event). Der Organizer legt
  // sie über einen Kalender an, die Anmeldeseite zeigt sie als Kalender.
  // Piggyback _subEventCalendar.
  const [subEventCalendar, setSubEventCalendar] = React.useState<boolean>(!!(editEvent && editEvent.subEventCalendar));
  // v29.22: Zum Löschen vorgemerkte, GESPEICHERTE Termine (Drafts mit dbId,
  // per Kalender-Klick oder X abgewählt). Sie bleiben hier geparkt statt
  // einfach zu verschwinden: Der Kalender zeigt sie ORANGE („wird beim
  // Speichern gelöscht"), ein erneuter Klick holt den Draft mitsamt allen
  // Einstellungen zurück. Vorher war ein abgewählter gespeicherter Tag
  // optisch nicht von „nie dagewesen" zu unterscheiden — und durch die
  // keyboard-selected-Färbung des DatePickers sah er sogar weiter grün aus,
  // als hätte das Abwählen nicht funktioniert. Die Speicher-Mechanik ändert
  // sich nicht: Nicht in subEvents = nicht in keptDbIds = wird beim Save
  // (nach der bestehenden Rückfrage) gelöscht.
  const [removedSavedSubs, setRemovedSavedSubs] = React.useState<SubEventDraft[]>([]);
  // v29.22: Die Terminliste unter dem Kalender ist standardmäßig EINGEKLAPPT
  // — bei 20 Terminen war sie eine Bildschirmseite Wiederholung dessen, was
  // der Kalender schon zeigt. Aufklappen nur bei Bedarf (Bearbeiten/Details).
  const [terminListOpen, setTerminListOpen] = React.useState(false);
  // v28.97: Genau EIN Sub-Event waehlbar statt beliebig vieler.
  const [subEventSingleChoice, setSubEventSingleChoice] = React.useState<boolean>(!!(editEvent && editEvent.subEventSingleChoice));
  // v30.61: Gebündelte Kommunikation (Mail / Kalender / QR getrennt schaltbar).
  // Gelesen aus den Overrides der Klammer — siehe utils/bundledComm.
  const [bundledComm, setBundledComm] = React.useState<BundledComm>(() => bundledCommOf(editEvent));
  // v28.10: Seitenverhältnis des Wizard-Bilds — die Banner-Option ist nur für
  // Querformat-Fotos sinnvoll und wird nur dann angeboten (Ratio >= 1.2).
  const [wizardImgAspect, setWizardImgAspect] = React.useState<number | null>(null);
  React.useEffect(() => {
    if (!imagePreview) { setWizardImgAspect(null); return; }
    let cancelled = false;
    const img = new Image();
    img.onload = () => { if (!cancelled && img.naturalHeight > 0) setWizardImgAspect(img.naturalWidth / img.naturalHeight); };
    img.src = imagePreview;
    return () => { cancelled = true; };
  }, [imagePreview]);
  React.useEffect(() => {
    // Nicht-Querformat (z.B. nach Kreis-Zuschnitt) → Banner-Flag zurücknehmen,
    // sonst bliebe ein unsichtbar gesetztes _imageBanner am Event hängen.
    if (wizardImgAspect != null && wizardImgAspect < 1.2 && imageBanner) setImageBanner(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [wizardImgAspect]);
  // v11.20: Re-sync useEffect aus v11.19 wieder rausgenommen — der hat
  // den Wizard-State mit stale-editEvent-Daten überschrieben (re-sync 2
  // mit helpText="" wurde im Maintainer-DevTools beobachtet, obwohl SP
  // nachweislich helpText="Test123" hatte). Das Aufrufen von
  // setCustomFields aus dem Effect heraus war zu fragil. Stattdessen
  // verlassen wir uns wieder auf den useState-Initializer + zusätzlich
  // ein detaillierteres Save-Log um zu sehen was *wirklich* an SP geht.
  const [customFields, setCustomFields] = React.useState<CustomFieldInput[]>(
    editEvent ? editEvent.eventSpecificFields.map(f => ({
      id: f.id, label: f.label, type: f.type, required: f.required,
      options: f.options ? [...f.options] : [], visible: true,
      ...(f.multi ? { multi: true } : {}),
      ...(f.defaultValue ? { defaultValue: f.defaultValue } : {}),
      ...(f.optionCategories && f.optionCategories.length > 0 ? { optionCategories: [...f.optionCategories] } : {}),
      ...(f.prefilterLabel ? { prefilterLabel: f.prefilterLabel } : {}),
      ...(f.helpText ? { helpText: f.helpText } : {}),
      ...(f.helpTextStyle === 'inline' ? { helpTextStyle: 'inline' as const } : {}),
      ...(f.showIf ? { showIf: { fieldId: f.showIf.fieldId, values: [...f.showIf.values] } } : {}),
      ...(f.onlyForGroup ? { onlyForGroup: f.onlyForGroup } : {}),
      // v11.94: confirmLabel beim Edit-Mount mit-übernehmen.
      ...(f.confirmLabel ? { confirmLabel: f.confirmLabel } : {}),
      // v17.20: EN-Varianten beim Edit-Mount mit-übernehmen.
      ...(f.labelEn ? { labelEn: f.labelEn } : {}),
      ...(f.helpTextEn ? { helpTextEn: f.helpTextEn } : {}),
      ...(f.confirmLabelEn ? { confirmLabelEn: f.confirmLabelEn } : {}),
      ...(f.optionsEn && f.optionsEn.length > 0 ? { optionsEn: [...f.optionsEn] } : {}),
      ...(f.externalLinks && f.externalLinks.length > 0 ? { externalLinks: f.externalLinks.map(x => ({ ...x })) } : {}),
      // v18.41: CC-bei-Mail-Flag beim Edit-Mount mit-übernehmen.
      ...(f.ccOnEmails ? { ccOnEmails: true } : {}),
      // v26.60: abgeschaltete Roommate-Benachrichtigung mit-übernehmen.
      ...(f.notifyRoommate === false ? { notifyRoommate: false } : {}),
      // v29.40: Verteiler-Begrenzung des Personen-Feldes mit-übernehmen.
      ...(f.audienceOnly ? { audienceOnly: true } : {}),
      // v29.20 (Audit A3): withTime (v24.25) und die daterange-Grenzen
      // (v28.63) fehlten in DIESEM Mapper — serializeCustomFields schreibt
      // nur, was im Draft steht, also entfernte jeder Edit-Save die
      // Einstellungen still: „Datum + Uhrzeit" wurde zum reinen Datum, das
      // buchbare Übernachtungs-Fenster verschwand. Die historische
      // Drop-Klasse (v11.21/v18.20/v19.20) saß diesmal im Lade-Pfad.
      ...(f.withTime ? { withTime: true } : {}),
      ...(f.rangeStart ? { rangeStart: f.rangeStart } : {}),
      ...(f.rangeEnd ? { rangeEnd: f.rangeEnd } : {}),
      ...(typeof f.maxNights === 'number' && f.maxNights > 0 ? { maxNights: f.maxNights } : {}),
    })) : []
  );
  const [outlookBody, setOutlookBody] = React.useState(editEvent ? reinsertOrganizerPlaceholder(stripOutlookWrapper(editEvent.outlookBody || ''), editEvent.organizers || []) : '');
  // Outlook-Termin-Header: beide Überschriften sind pro Event editierbar.
  // Default: eventTitle + formatiertes Startdatum. Parsed aus bestehendem
  // OutlookBody, falls der User sie schon angepasst hat.
  const [outlookHeading, setOutlookHeading] = React.useState(() => {
    if (editEvent) {
      const p = parseOutlookHeadings(editEvent.outlookBody || '');
      if (p.heading) return p.heading;
    }
    return editEvent ? (editEvent.title || '') : '';
  });
  const [outlookSubheading, setOutlookSubheading] = React.useState(() => {
    if (editEvent) {
      const p = parseOutlookHeadings(editEvent.outlookBody || '');
      if (p.subheading && p.subheading !== 'Event Details') return p.subheading;
    }
    return '';
  });
  // v18.42: Betreff des Outlook-Termins (leer = Event-Titel). Per-Tab gespiegelt
  // wie outlookHeading; persistiert in der DEX_Events-Spalte OutlookSubject.
  const [outlookSubject, setOutlookSubject] = React.useState<string>(editEvent?.outlookSubject || '');
  // v18.44: abweichendes Outlook-Datum (Top-Level). Leer = Event-Start/-Ende.
  // Als ISO gespeichert (wie Sub-Event-Datum); DatePicker konvertiert via isoToLocal.
  // v29.52: Ganztägiger Termin (Hauptevent). Die Krücke 00:00–23:59 blockiert in
  // Outlook den Tag als normalen Termin statt als Ganztags-Eintrag im Kopf —
  // der Unterschied fällt erst im Kalender des Teilnehmers auf.
  const [allDay, setAllDay] = React.useState<boolean>(!!(editEvent && editEvent.allDay));
  // v29.54: Der Termin blockiert den Kalender (Outlook `showAs`). Default ist
  // beschäftigt; bei Ganztags-Terminen ist das oft nicht gewollt, weil dann
  // ein ganzer Arbeitstag als belegt gilt.
  const [showAsFree, setShowAsFree] = React.useState<boolean>(!!(editEvent && editEvent.showAsFree));
  // v29.55: Bekommen die Organizer den Outlook-Termin JEDES Sub-Events? Der
  // Flow setzt requiredAttendees aus OrganizerEmail, und die steht auf jeder
  // Sub-Event-Zeile — bei 21 Tagen sind das 21 Blocker im Kalender, für Tage
  // ohne eigene Buchung. Positiv im UI, negativ gespeichert (skipOrganizerInvite).
  // Die Einstellung gilt event-weit: Klammer und alle Sub-Events bekommen
  // denselben Wert. Bestandsevents kommen mit false an und bleiben unverändert.
  const [orgGetsSubInvites, setOrgGetsSubInvites] = React.useState<boolean>(
    editEvent ? !editEvent.skipOrganizerInvite : true,
  );
  // Hat der Organizer die Entscheidung selbst getroffen? Dann nie überschreiben.
  const orgInvitesTouchedRef = React.useRef<boolean>(!!editEvent);
  // v29.56: Stand beim Oeffnen — daraus leitet sich beim Speichern ab, ob die
  // Organizer an BESTEHENDEN Outlook-Terminen nachträglich an- oder
  // abgemeldet werden müssen. SkipOrganizerInvite wirkt nämlich nur beim
  // ANLEGEN (requiredAttendees); ein bestehender Termin behält seine
  // Teilnehmerliste, bis jemand Einladen/Ausladen queued.
  const initialOrgGetsSubInvitesRef = React.useRef<boolean>(editEvent ? !editEvent.skipOrganizerInvite : true);
  const [outlookStartOverride, setOutlookStartOverride] = React.useState<string>(editEvent?.outlookStart || '');
  const [outlookEndOverride, setOutlookEndOverride] = React.useState<string>(editEvent?.outlookEnd || '');
  // Modal-State für den HTML-Editor (Outlook-Body + E-Mail-Templates)
  const [htmlEditorOpen, setHtmlEditorOpen] = React.useState(false);
  const [htmlEditorMode, setHtmlEditorMode] = React.useState<'outlook' | 'email' | 'description'>('outlook');
  const [htmlEditorTemplateType, setHtmlEditorTemplateType] = React.useState<string>('');
  const [emailLanguage, setEmailLanguage] = React.useState(
    editEvent
      ? (editEvent.emailLanguage || (locale === 'de' ? 'DE' : 'EN'))
      : (locale === 'de' ? 'DE' : 'EN')
  );
  const [disableEmails, setDisableEmails] = React.useState(editEvent ? !!editEvent.disableEmails : false);
  // v19.21: granulare Sub-Schalter unter dem Master „Bestätigungs-E-Mails":
  // einzeln die Anmelde- bzw. Abmelde-Bestätigung abschaltbar (Top-Level-Event).
  const [disableRegistrationEmail, setDisableRegistrationEmail] = React.useState(editEvent ? !!editEvent.disableRegistrationEmail : false);
  const [disableCancellationEmail, setDisableCancellationEmail] = React.useState(editEvent ? !!editEvent.disableCancellationEmail : false);
  // v19.23: Outlook-Absage = automatische Abmeldung vom Event (Flow-getrieben,
  // Top-Level-Event). Persistiert als DEX_Events-Spalte; der
  // DEX_OutlookDeclineHandler-Flow liest die Spalte und meldet die Person ab.
  const [autoDeregisterOnDecline, setAutoDeregisterOnDecline] = React.useState(editEvent ? !!editEvent.autoDeregisterOnDecline : false);
  const [inactiveHandling, setInactiveHandling] = React.useState<'notify' | 'autoderegister'>(editEvent && editEvent.inactiveHandling === 'autoderegister' ? 'autoderegister' : 'notify');
  const [disableOutlook, setDisableOutlook] = React.useState(editEvent ? !!editEvent.disableOutlook : false);
  // v14.4: Acknowledgement, dass bei Top-Level-Kommunikation = AUS die
  // Teilnehmer sich für mindestens ein Sub-Event anmelden müssen. Vorausgewählt
  // für Events, die schon mit deaktivierter Kommunikation gespeichert sind
  // (alter Lauf ist bereits durch die Gate-Logik durchgekommen). Bei neuen
  // Events / frisch umgeschaltetem Toggle bleibt der Haken aus, der Save
  // ist dann blockiert bis bestätigt.
  const [mainCommDisabledAck, setMainCommDisabledAck] = React.useState<boolean>(
    !!editEvent && (!!editEvent.disableEmails || !!editEvent.disableOutlook),
  );
  // v14.5: Toggle „Anmeldung für mindestens ein Sub-Event verpflichtend".
  // Wird im RegistrationForm erzwungen — der Submit-Button blockiert, bis
  // der Teilnehmer ein Sub-Event angehakt hat. Sinnvoll wenn die Haupt-
  // Event-Kommunikation aus ist und alle Mails/Outlook-Termine nur über
  // die Sub-Events laufen.
  // v15.3: setRequireSubEventSelection wird nicht mehr direkt von der UI
  // aufgerufen — der Flag wird beim Save aus dem subEventsOnlyMode-Toggle
  // in Schritt 2 abgeleitet. State bleibt als Read-only für die Save-
  // Logik (siehe handleSubmit).
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const [requireSubEventSelection, setRequireSubEventSelection] = React.useState<boolean>(
    !!editEvent && !!editEvent.requireSubEventSelection,
  );
  // v14.8: „Nur Sub-Events"-Modus. Wenn true, ist die Hauptevent-Anmeldung im
  // Teilnehmerformular ausgeblendet. Impliziert requireSubEventSelection=true.
  const [subEventsOnlyMode, setSubEventsOnlyMode] = React.useState<boolean>(
    !!editEvent && !!editEvent.subEventsOnlyMode,
  );
  // v14.8: Organizer-konfigurierbarer Begriff für die untergeordneten Events
  // (Standard „Sub-Event" / „Sub-Events", alternativ Workshop / Session etc.).
  // v15.9: separater `customTermMode`-Flag, damit „Eigene Bezeichnung…"
  // im Dropdown auch dann angeklebt bleibt, wenn beide Inputs noch leer
  // sind (sonst kippt die Heuristik unten auf 'subevent' zurück und die
  // Custom-Inputs verschwinden bevor der User tippen kann).
  const [customTermMode, setCustomTermMode] = React.useState<boolean>(false);
  // v29.60: Grammatisches Geschlecht der Bezeichnung. '' = wie bisher raten.
  // Gebraucht wird der AKKUSATIV („wähle mindestens einen Office-Tag aus"),
  // und den kann man aus dem Wort nicht ableiten — deshalb gefragt statt geraten.
  const [childGender, setChildGender] = React.useState<'' | 'm' | 'f' | 'n'>(
    (editEvent && editEvent.childEventTermGender) || '',
  );
  const [childTermSingular, setChildTermSingular] = React.useState<string>(
    (editEvent && editEvent.childEventTermSingular) || '',
  );
  const [childTermPlural, setChildTermPlural] = React.useState<string>(
    (editEvent && editEvent.childEventTermPlural) || '',
  );
  // v8.5: Organizer-BCC-Modi (Anmeldung + Abmeldung).
  const [notifyOrgRegisterMode, setNotifyOrgRegisterMode] = React.useState<'never' | 'always' | 'fromDate'>(
    editEvent ? (editEvent.notifyOrgRegisterMode || 'never') : 'never'
  );
  // v29.19: Wie alle anderen Datumsfelder über isoToLocal laden — der rohe
  // SP-Wert ist UTC-ISO mit „Z", und berlinLocalToUtcIso beim Save hängt ein
  // weiteres „Z" an → ungültig → es wurde '' in die Spalte geschrieben. Wer
  // das Feld beim Edit nicht neu anfasste, verlor sein „BCC ab"-Datum still;
  // der Modus blieb auf FromDate stehen und die Organizer-Kopie feuerte nie.
  const [notifyOrgRegisterFromDate, setNotifyOrgRegisterFromDate] = React.useState<string>(
    editEvent && editEvent.notifyOrgRegisterFromDate ? isoToLocal(editEvent.notifyOrgRegisterFromDate) : ''
  );
  const [notifyOrgCancelMode, setNotifyOrgCancelMode] = React.useState<'never' | 'always' | 'afterDeadline'>(
    // v10.17+: Default für neue Events ist 'afterDeadline' (Erst nach der
    // letzten Abmeldemöglichkeit) — sonst flutet jede Stornierung den
    // Organizer-Posteingang. User-Wunsch.
    editEvent ? (editEvent.notifyOrgCancelMode || 'never') : 'afterDeadline'
  );
  // v8.6: Exclude-Liste — explizit ausgeschlossene User (überschreiben den
  // Sichtbarkeits-Filter). UI: Modal "Sichtbare Personen anzeigen".
  const [excludedUsers, setExcludedUsers] = React.useState<string[]>(
    editEvent ? (editEvent.excludedUsers || []) : []
  );
  // v11.88: Demo-Auswahl-Modal — der „Demo"-Button öffnet einen Dialog
  // mit vier Vorlagen-Karten (Standard, Mit Gruppen, Mit Sub-Event,
  // Mit Sub-Event + Team). Klick auf eine Karte füllt das Formular
  // mit der jeweiligen Variante und schliesst das Modal.
  const [showDemoVariantModal, setShowDemoVariantModal] = React.useState<boolean>(false);
  // v24.9 (E): „Eigenes Event als Vorlage" — aufklappbare Kachelgalerie der
  // bisherigen Events des Organizers; Auswahl lädt Einstellungen + Bild.
  const [showTemplatePicker, setShowTemplatePicker] = React.useState<boolean>(false);
  const [templateLoadingId, setTemplateLoadingId] = React.useState<string | null>(null);
  // v17.21: Modal nach erfolgreichem Speichern — fragt den Organizer, ob er
  // eine A4-Zusammenfassung des Events herunterladen möchte. Pending-Payload
  // hält die Info für den `dex-event-submit-success`-Dispatch, der erst
  // gefeuert wird, wenn der User im Modal eine Auswahl getroffen hat.
  const [showSummaryModal, setShowSummaryModal] = React.useState<boolean>(false);
  const [pendingSuccessDispatch, setPendingSuccessDispatch] = React.useState<{
    title: string; eventId: string; type: 'create' | 'update';
  } | null>(null);
  // v17.22: Unmount-Safety. Der Success-Dispatch (dex-event-submit-success,
  // treibt Erfolgs-Banner + Auto-Navigation in DexEventPlatform) läuft erst,
  // wenn der User im Summary-Modal eine Auswahl trifft. Verlässt er den
  // Wizard vorher (Header-Navigation, Browser-Back, Tab-Eviction), würde der
  // Dispatch sonst verloren gehen — Folge: kein Banner, kein Redirect, User
  // denkt der Save sei fehlgeschlagen. Dieser Ref + Cleanup-Effect feuert den
  // Dispatch beim Unmount nach, falls er noch aussteht.
  const pendingSuccessDispatchRef = React.useRef<{ title: string; eventId: string; type: 'create' | 'update' } | null>(null);
  React.useEffect(() => {
    return () => {
      const pending = pendingSuccessDispatchRef.current;
      if (pending) {
        pendingSuccessDispatchRef.current = null;
        try {
          window.dispatchEvent(new CustomEvent('dex-event-submit-success', { detail: pending }));
        } catch { /* */ }
      }
    };
  }, []);
  // v19.x: Der gesamte Ausschluss-Modal-State (resolved Members, Suche,
  // Tabellen-Filter, Sortierung, Pagination) ist nach <AudiencePicker>
  // gewandert. Hier bleibt nur die persistierte `excludedUsers`-Liste (oben),
  // die als Prop in den Picker durchgereicht wird.
  // v9.16: neue Events starten standardmäßig als Test-Event — der Organizer
  // kann sich erst alles in Ruhe anschauen, das Test-Team probiert die
  // Anmeldung durch, und erst wenn alles passt wird der Schalter rausgenommen.

  return {
    allDay, audience, autoDeregisterOnDecline, bundledComm, childGender, childTermPlural,
    childTermSingular, customFields, customTermMode, description, disableCancellationEmail, disableEmails,
    disableOutlook, disableRegistrationEmail, emailLanguage, emailLogoFromPhoto, endDate, eventImageUrl,
    excludedUsers, filterMode, htmlEditorMode, htmlEditorOpen, htmlEditorTemplateType, imageBanner,
    imageDisplay, imageDisplayOpen, imageEditOpen, imageFile, imageOrigAspect, imageOrigFile,
    imagePreview, inactiveHandling, initialOrgGetsSubInvitesRef, klammerDeadline, lastDeregisterDate, locationFilter,
    logoCropTarget, mainCommDisabledAck, maxParticipants, noCancelAfterDeadline, noDescription, notifyOrgCancelMode,
    notifyOrgRegisterFromDate, notifyOrgRegisterMode, onlineMeetingMode, orgGetsSubInvites, orgInvitesTouchedRef, outlookBody,
    outlookEndOverride, outlookHeading, outlookLocationOverride, outlookLogoFromPhoto, outlookStartOverride, outlookSubheading,
    outlookSubject, pendingSuccessDispatch, pendingSuccessDispatchRef, registrationDeadline, removedSavedSubs, requireSubEventSelection,
    setAllDay, setAudience, setAutoDeregisterOnDecline, setBundledComm, setChildGender, setChildTermPlural,
    setChildTermSingular, setCustomFields, setCustomTermMode, setDescription, setDisableCancellationEmail, setDisableEmails,
    setDisableOutlook, setDisableRegistrationEmail, setEmailLanguage, setEmailLogoFromPhoto, setEndDate, setEventImageUrl,
    setExcludedUsers, setFilterMode, setHtmlEditorMode, setHtmlEditorOpen, setHtmlEditorTemplateType, setImageBanner,
    setImageDisplay, setImageDisplayOpen, setImageEditOpen, setImageFile, setImageOrigAspect, setImageOrigFile,
    setImagePreview, setInactiveHandling, setKlammerDeadline, setLastDeregisterDate, setLocationFilter, setLogoCropTarget,
    setMainCommDisabledAck, setMaxParticipants, setNoCancelAfterDeadline, setNoDescription, setNotifyOrgCancelMode, setNotifyOrgRegisterFromDate,
    setNotifyOrgRegisterMode, setOnlineMeetingMode, setOrgGetsSubInvites, setOutlookBody, setOutlookEndOverride, setOutlookHeading,
    setOutlookLocationOverride, setOutlookLogoFromPhoto, setOutlookStartOverride, setOutlookSubheading, setOutlookSubject, setPendingSuccessDispatch,
    setRegistrationDeadline, setRemovedSavedSubs, setRequireSubEventSelection, setShowAsFree, setShowDemoVariantModal, setShowSummaryModal,
    setShowTemplatePicker, setStartDate, setSubEventCalendar, setSubEventSingleChoice, setSubEventsOnlyMode, setSubImageCropIdx,
    setTeamsLink, setTemplateLoadingId, setTerminListOpen, setUnlimitedParticipants, setUserCancelAllowed, setWaitlistEnabled,
    showAsFree, showDemoVariantModal, showSummaryModal, showTemplatePicker, startDate, storedEventType,
    subEventCalendar, subEventSingleChoice, subEventsOnlyMode, subImageCropIdx, teamsLink, templateLoadingId,
    terminListOpen, unlimitedParticipants, userCancelAllowed, waitlistEnabled, wizardImgAspect,
  };
}
