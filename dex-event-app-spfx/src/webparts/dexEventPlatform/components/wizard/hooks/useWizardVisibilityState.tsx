/* useWizardVisibilityState — aus EventCreationPage.tsx ausgelagert (Zeilen 1024-1526 des
 * damaligen Stands). Der Bereich ist zeichengleich uebernommen, inklusive der
 * Reihenfolge seiner Hook-Aufrufe: der Aufruf dieses Hooks steht in der Seite
 * an genau der Stelle, an der der Bereich vorher stand. Schwerpunkt: Sichtbarkeit, Entwurfsstatus, Bildzuschnitt.
 *
 * Die Grenze ist mechanisch gezogen (zusammenhaengender Bereich), nicht
 * thematisch — der Name beschreibt den Schwerpunkt, nicht eine reine Trennung. */
import * as React from 'react';
import { EmailOverrideEntry } from '../../wizard/emailOverrideEntry';
import { readOutlookLogo } from '../../wizard/wizardHelpers';
import { compressImage } from '../../../utils/imageCompress';
import { applyEventPhotoToLogoImpl } from '../../wizard/logic/wizardMisc';
import { renderHeaderSizeControlImpl } from '../../wizard/logic/wizardRenderHelpers';
import { AgendaItem } from '../../../types';
import { SubEventDraft } from '../../wizard/wizardTypes';
import { parseOutlookHeadings, stripOutlookWrapper } from '../../../services/EmailTemplates';
import { CustomFieldInput } from '../../wizard/customFieldInput';

export interface UseWizardVisibilityStateCtx {
  childEventsOf: (parentEventId: string) => import("../../../types/index").DeloitteEvent[];
  editEvent: import("../../../types/index").DeloitteEvent;
  imageFile: File;
  imageOrigFile: File;
  imagePreview: string;
  isDe: boolean;
  locale: import("../../../context/LanguageContext").Locale;
  onlineMeetingMode: "auto" | "none" | "own";
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  teamsLink: string;
}

export function useWizardVisibilityState(ctx: UseWizardVisibilityStateCtx) {
  const { childEventsOf, editEvent, imageFile, imageOrigFile, imagePreview, isDe, locale, onlineMeetingMode, showAlert, teamsLink } = ctx;
  const [isFictive, setIsFictive] = React.useState(editEvent ? !!editEvent.isFictive : true);
  // v18.9: Organizer-Anzeige (Chips mit Name + Foto) auf Anmelde-Seite +
  // „Meine Events" ausblenden. Rein visuell — Rechte/Mails unberührt.
  const [hideOrganizer, setHideOrganizer] = React.useState(editEvent ? !!editEvent.hideOrganizer : false);
  // v24.15: Wenn „Organizer ausblenden" an ist UND es mehrere Organizer gibt,
  // kann der Organizer stattdessen NUR EINZELNE ausblenden (Piggyback
  // `_hideOrgIndividual`). false = alle ausblenden; true = nur die angeklickten.
  const [hideOrganizerIndividualOnly, setHideOrganizerIndividualOnly] = React.useState(editEvent ? !!editEvent.hideOrganizerIndividualOnly : false);
  // v23.6: „Assistenzen sehen das Event generell" (Piggyback _assistantsCanSee).
  // Wenn aktiv, sehen Personen mit dem Job-Title „Assistenz" das Event auch
  // dann, wenn Standort-/Verteiler-Filter sie sonst ausschließen würden —
  // damit sie stellvertretend (z.B. für einen Partner) anmelden können.
  const [assistantsCanSee, setAssistantsCanSee] = React.useState(editEvent ? !!editEvent.assistantsCanSee : false);
  // v23.25: Organizer auf der Anmeldeseite groß (Foto + Mail + Rolle direkt
  // sichtbar) statt klein als Chip mit Hover (Piggyback _organizerDisplayLarge).
  const [organizerDisplayLarge, setOrganizerDisplayLarge] = React.useState(editEvent ? !!editEvent.organizerDisplayLarge : false);
  // v24.8 (J): EINZELNE Organizer von der Anzeige ausnehmen (Klick auf den Chip).
  // Sie behalten alle Rechte/Mails — sie werden nur nicht als Ansprechpartner
  // auf der Anmelde-Seite gezeigt. Piggyback `_hiddenOrganizers` (E-Mails, lc).
  const [hiddenOrganizerEmails, setHiddenOrganizerEmails] = React.useState<string[]>(
    editEvent && editEvent.hiddenOrganizerEmails ? editEvent.hiddenOrganizerEmails.map(e => (e || '').toLowerCase()).filter(Boolean) : []
  );
  const toggleOrganizerHidden = (email: string): void => {
    const lc = (email || '').toLowerCase();
    if (!lc) return;
    setHiddenOrganizerEmails(prev => prev.indexOf(lc) >= 0 ? prev.filter(x => x !== lc) : [...prev, lc]);
  };
  // Nur im Edit-Modus: standardmäßig wird der Outlook-Termin NICHT angefasst,
  // damit bei kleinen Änderungen (z.B. Description) nicht unnötig eine
  // "Updated meeting"-Benachrichtigung an alle Teilnehmer geht. Der Organizer
  // muss die Checkbox aktiv setzen wenn er möchte dass Titel/Start/Ende im
  // Outlook-Termin aktualisiert werden.
  const [triggerOutlookUpdate, setTriggerOutlookUpdate] = React.useState(false);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [emailTemplates, setEmailTemplates] = React.useState<Array<{ id: number; templateType: string; language: string; subject: string; heading: string; headingColor: string; bodyHtml: string }>>([]);
  const [emailTemplateOverrides, setEmailTemplateOverrides] = React.useState<Record<string, EmailOverrideEntry>>(
    editEvent?.emailTemplateOverrides ? (() => {
      try {
        const parsed = JSON.parse(editEvent.emailTemplateOverrides);
        // v11.39: Alle Piggyback-Keys rausstrippen — sie werden in separaten
        // States gehalten (emailLogoPreview, outlookLogoPreview, testTeamEmails
        // etc.) und beim Speichern frisch dazugemerged. Wenn sie hier
        // mitgeschleppt werden, überschreibt der spread `...emailTemplateOverrides`
        // am Ende von handleSubmit die frisch berechneten Werte und das
        // Entfernen z.B. eines Test-Team-Mitglieds bleibt ohne Wirkung.
        const {
          _eventLogo, _outlookLogo, _outlookLogoSameAsMail, _b2run,
          _qrScanners, _coOrganizers, _testTeam,
          _splitDisplayOrderReversed,
          _requireSubEventSelection,
          _subEventsOnlyMode, _subEventsDisabled, _imageBanner, _childEventTerm,
          _inheritFlags, _hideOrganizer, _headerImageLayout,
          // v22.78/v23.6: diese Piggyback-Keys MÜSSEN ebenfalls gestrippt
          // werden — sonst überschreibt der stale Wert aus dem geladenen Blob
          // beim Edit-Save (letzter Spread `...topOverrides`) das frisch
          // berechnete Flag, d.h. Abwählen bliebe ohne Wirkung.
          _teamTerm, _teamMembersCannotCreate, _assistantsCanSee, _previewBeforeActive, _imageDisplay,
          _organizerDisplayLarge, _hiddenOrganizers, _hideOrgIndividual, _mainEventLabel,
          _imageOrigUrl, _klammerDeadline,
          // v28.39: Hotel-Planung wird ausschliesslich im Organizer Center gepflegt.
          // Stripping verhindert, dass ein parallel offener Wizard beim Speichern
          // einen veralteten Stand zurückschreibt.
          _hotels, _hotelStays, _hotelVisible, _hotelRules,
          // v28.79: „Keine Beschreibung nutzen"-Flag (s. noDescriptionConfig).
          _noDescription,
          // v28.91: Kalender-Modus der Sub-Events (s. subEventCalendarConfig).
          _subEventCalendar, _subEventSingleChoice,
          // v29.25: Abmelde-Sperren (s. userCancelAllowed / noCancelAfterDeadline).
          _noSelfCancel, _noCancelAfterDeadline,
          // v29.38: Teams-Link (s. teamsLinkConfig).
          _teamsLink,
          // v29.66: Abrechnungs-Piggyback (F&A-Pilot) — lebt in eigenen States.
          _billing,
          // v29.67: Freischalt-Regel der Kalender-Termine — eigene States.
          _subEventOpenRule,
          // v29.75: Sichtbarkeit-für-alle-Sub-Events-Haken — eigener State.
          _visAllSubs,
          // v29.76: Rollierende Fristen der Kalender-Termine — eigene States.
          _subDeadlineRule,
          // v30.61: Gebündelte Kommunikation — eigene States (s. bundledComm).
          _commBundledMail, _commBundledOutlook, _commBundledQr,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          ...rest
        } = parsed as Record<string, unknown>;
        // Variablen nur destrukturiert, um sie aus `rest` zu entfernen.
        void _eventLogo; void _outlookLogo; void _outlookLogoSameAsMail; void _b2run;
        void _billing; void _subEventOpenRule; void _visAllSubs; void _subDeadlineRule;
        void _qrScanners; void _coOrganizers; void _testTeam;
        void _splitDisplayOrderReversed; void _requireSubEventSelection;
        void _subEventsOnlyMode; void _subEventsDisabled; void _imageBanner; void _childEventTerm;
        void _inheritFlags; void _hideOrganizer; void _headerImageLayout;
        void _teamTerm; void _teamMembersCannotCreate; void _assistantsCanSee; void _previewBeforeActive; void _imageDisplay;
        void _organizerDisplayLarge; void _hiddenOrganizers; void _hideOrgIndividual; void _mainEventLabel;
        void _imageOrigUrl; void _klammerDeadline; void _noDescription;
        void _subEventCalendar; void _subEventSingleChoice;
        void _noSelfCancel; void _noCancelAfterDeadline; void _teamsLink;
        void _hotels; void _hotelStays; void _hotelVisible; void _hotelRules;
        void _commBundledMail; void _commBundledOutlook; void _commBundledQr;
        return rest as Record<string, EmailOverrideEntry>;
      } catch { return {}; }
    })() : {}
  );
  // editingTemplate state entfällt seit Modal-Migration v4.7.0
  // Custom Event-Logo für E-Mails (ersetzt das DEX-Orb in E-Mails).
  const [emailLogoPreview, setEmailLogoPreview] = React.useState<string>(() => {
    if (!editEvent?.emailTemplateOverrides) return '';
    try { const o = JSON.parse(editEvent.emailTemplateOverrides); return o._eventLogo || ''; } catch { return ''; }
  });
  // Custom Event-Logo für Outlook-Termin (ersetzt das DEX-Orb im Termin-Body).
  // Separat vom Mail-Logo, damit man z.B. in Mails das neutrale DEX-Logo lassen
  // und im Outlook-Termin ein event-spezifisches Bild anzeigen kann.
  const [outlookLogoPreview, setOutlookLogoPreview] = React.useState(() => {
    if (!editEvent?.emailTemplateOverrides) return '';
    try { return readOutlookLogo(JSON.parse(editEvent.emailTemplateOverrides)); } catch { return ''; }
  });
  // v18.73: Header-Bild (Event-Bild = {{ORB_URL}}) Größe + Innenabstand pro
  // Event. Gilt für Mail- UND Outlook-Termin-Kopf. Persistiert als Piggyback
  // `_headerImageLayout` in EmailTemplateOverrides. Default = bisheriges
  // Layout (Breite 180, Innenabstand 30/30).
  // v29.29: NEUE Events starten mit dem Vollbild-Kopf (Bild über die ganze
  // Mailbreite) — das ist der Regelfall, das kleine zentrierte Bild war eher
  // die Ausnahme. Bestehende Events ohne gespeichertes Layout behalten
  // bewusst 180/30/30: Ihre Mails sähen sonst nach dem nächsten Speichern
  // ungefragt anders aus.
  const [headerImageLayout, setHeaderImageLayout] = React.useState<{ width: number; paddingV: number; paddingH: number }>(() => {
    const legacyDef = { width: 180, paddingV: 30, paddingH: 30 };
    const fullWidth = { width: 600, paddingV: 0, paddingH: 0 };
    if (!editEvent) return fullWidth;
    if (!editEvent.emailTemplateOverrides) return legacyDef;
    try {
      const o = JSON.parse(editEvent.emailTemplateOverrides);
      const il = o._headerImageLayout || {};
      return {
        width: typeof il.width === 'number' && il.width > 0 ? il.width : 180,
        paddingV: typeof il.paddingV === 'number' && il.paddingV >= 0 ? il.paddingV : 30,
        paddingH: typeof il.paddingH === 'number' && il.paddingH >= 0 ? il.paddingH : 30,
      };
    } catch { return legacyDef; }
  });
  // v19.20: Snapshot des initialen Header-Bild-Layouts (Breite/Innenabstand)
  // beim Edit-Mount. Eine reine Layout-Änderung verändert NICHT den rohen
  // Outlook-Body-Text (das Layout wird erst beim Wrappen via buildOutlookBody
  // angewendet) — der Outlook-Änderungs-Detektor hätte sie deshalb übersehen.
  // Wir vergleichen das aktuelle Layout gegen diesen Snapshot, damit eine
  // Größen-/Abstands-Änderung das „Outlook-Termin aktualisieren?"-Modal genauso
  // öffnet wie eine Textänderung. useRef fixiert den Wert beim ersten Render.
  const initialHeaderImageLayoutRef = React.useRef<{ width: number; paddingV: number; paddingH: number }>(headerImageLayout);
  // v29.38: Gleiche Mechanik für den Teams-Link — er steht nicht im rohen
  // Termin-Text (er wird erst beim Wrappen angehängt), ändert den Termin aber
  // sichtbar. Ohne Snapshot bliebe eine reine Link-Änderung für den
  // Update-Detektor unsichtbar und der Termin behielte den alten Stand.
  const effTeamsLink = (): string => (onlineMeetingMode === 'own' ? teamsLink.trim() : '');
  /**
   * Link, der in den TERMIN-BODY wandert — nicht derselbe wie `effTeamsLink()`,
   * der nur den gespeicherten `_teamsLink` steuert.
   *
   * **v30.40: Im Modus „DEX erzeugt den Link" steht hier wieder nichts.** Der
   * Weg über die Marke `{{TEAMS_URL}}` (v30.27–v30.39) ist gescheitert, und
   * zwar erst im letzten möglichen Moment — im fertigen Termin beim Teilnehmer:
   *
   * 1. Die App schrieb `<a href="{{TEAMS_URL}}">…</a>` in den Body.
   * 2. Der Flow holte den Body per Graph und ersetzte die Marke durch die echte
   *    `joinUrl`. Der PATCH lief mit 200 durch.
   * 3. Im Termin stand danach `[https://teams.microsoft.com/l/meetup-join/…]An
   *    Microsoft-Teams-Besprechung teilnehmen` — der Anker war zu Text
   *    zerfallen. Vorher, mit der unersetzten Marke im href, war es noch ein
   *    Knopf.
   *
   * Beobachtet, nicht bewiesen: Exchange normalisiert den Body eines
   * Online-Meetings und lässt einen Anker auf die eigene joinUrl nicht stehen.
   * Was sich prüfen ließ, spricht dafür — die Degradierung trat exakt mit
   * unserem PATCH ein, an keiner früheren Stelle.
   *
   * Entscheidend ist aber nicht die Ursache, sondern dass der Kasten, den
   * Exchange unter die Karte hängt, ohnehin bleiben MUSS: Ihn zu entfernen
   * hieße, den Meeting-Blob aus dem Body zu werfen, und das deaktiviert die
   * Besprechung (Graph-Referenz `event: update`). Er trägt Join-Link,
   * Meeting-ID und Passcode — mehr, als der DEX-Block je hatte. Ein zweiter,
   * kaputter Link darüber macht den Termin nur schlechter.
   *
   * Für einen SELBST eingetragenen Link bleibt alles wie bisher: Da gibt es
   * keine Exchange-Normalisierung, der Block rendert sauber, und ohne ihn stünde
   * der Link nirgends.
   */
  const outlookTeamsLink = (): string => (
    onlineMeetingMode === 'own' ? teamsLink.trim() : ''
  );
  const initialTeamsLinkRef = React.useRef<string>(teamsLink);
  // v30.26: Der Modus zählt für den Outlook-Update-Detektor genauso wie der
  // Link selbst — ein Wechsel von „eigener Link" auf „DEX erzeugt" ändert
  // sowohl den Termin-Text (Link fällt weg) als auch den Termin-Typ
  // (isOnlineMeeting). Ohne diesen Vergleich bliebe der bestehende Termin
  // stehen, weil sich der reine teamsLink-String nicht bewegt hat.
  const initialOnlineMeetingModeRef = React.useRef<'none' | 'own' | 'auto'>(onlineMeetingMode);
  const onlineMeetingChanged = (): boolean =>
    onlineMeetingMode !== initialOnlineMeetingModeRef.current
    || effTeamsLink() !== (initialOnlineMeetingModeRef.current === 'own' ? initialTeamsLinkRef.current.trim() : '');
  // v18.73: Piggyback-Konfig für den Save (leer wenn alles auf Default steht —
  // dann wird der Key gar nicht geschrieben). Wird in Create- UND Edit-Pfad
  // sowie in die Sub-Event-Overrides gemerged.
  const headerImageLayoutConfig = (headerImageLayout.width !== 180 || headerImageLayout.paddingV !== 30 || headerImageLayout.paddingH !== 30)
    ? { _headerImageLayout: { width: headerImageLayout.width, paddingV: headerImageLayout.paddingV, paddingH: headerImageLayout.paddingH } }
    : {};
  /**
   * v28.29: Kopfbild-Layout für EINEN konkreten Outlook-/Mail-Body. Breite und
   * Innenabstand stellt der Organizer für sein FOTO ein („Volle Breite" = 600px).
   * Fällt ein Termin mangels eigenem Bild auf das Standard-DEX-Logo (Orb) zurück,
   * wurde dieses bisher ebenfalls 600px breit gerendert — in Outlook ein
   * bildschirmfüllender, unten abgeschnittener Orb. Ohne eigenes Bild deshalb
   * max. 180px mit Mindestabstand.
   */
  const headerLayoutFor = (logoB64: string): { imageWidth: number; imagePaddingV: number; imagePaddingH: number } => {
    const hasOwn = !!(logoB64 && logoB64.trim());
    return {
      imageWidth: hasOwn ? headerImageLayout.width : Math.min(headerImageLayout.width, 180),
      imagePaddingV: hasOwn ? headerImageLayout.paddingV : Math.max(headerImageLayout.paddingV, 20),
      imagePaddingH: hasOwn ? headerImageLayout.paddingH : Math.max(headerImageLayout.paddingH, 20),
    };
  };
  // v26.95: Das Event-Foto als Mail-/Outlook-Kopfbild übernehmen. Quelle ist der
  // frisch gewählte File (imageFile), sonst die Vorschau (Data-URL direkt, http-
  // URL bestehender Events wird geladen). In JEDEM Fall auf 600px komprimiert,
  // damit die Base64-Größe für die Mail-Pipeline handhabbar bleibt.
  const fileToBase64 = (file: File): Promise<string> =>
    new Promise<string>(resolve => { const r = new FileReader(); r.onload = e => resolve((e.target?.result as string) || ''); r.onerror = () => resolve(''); r.readAsDataURL(file); });
  // v28.10: Base64-Logos hart auf Mail-taugliche Größe bringen. Ungebremste
  // Logos landeten bis zu DREIMAL im selben Save-Payload (OutlookBody via
  // {{ORB_URL}}, EmailTemplateOverrides._eventLogo/_outlookLogo und
  // EmailImageBase64) und rissen das SharePoint-REST-Limit von 2 MB
  // („The request message is too big"). Ab ~400 KB wird auf 600px
  // runterskaliert; schlägt das fehl, bleibt der Originalwert.
  const shrinkLogoB64 = async (b64: string): Promise<string> => {
    // v28.31: Schwelle von 400 KB auf 200 KB und in Stufen verkleinern, bis das
    // Bild unter ~250 KB liegt. Dasselbe Base64 steckt DREIMAL im Save-Payload
    // (OutlookBody, _eventLogo, _outlookLogo) — bei 800 KB je Vorkommen reisst
    // ein einziges Event das SharePoint-Limit von 2 MB. Ausgabe immer als JPEG
    // auf weissem Grund (siehe compressImage): PNG-Fotos sind der Hauptgrund
    // für die Ausreisser.
    const TARGET = 250_000;
    if (!b64 || b64.indexOf('data:') !== 0 || b64.length <= 200_000) return b64;
    let best = b64;
    for (const w of [600, 480, 360]) {
      try {
        const resp = await fetch(best);
        const blob = await resp.blob();
        const f = new File([blob], 'logo.jpg', { type: blob.type || 'image/jpeg' });
        const out = await fileToBase64(await compressImage(f, w, 0.82, true));
        if (out && out.length < best.length) best = out;
        if (best.length <= TARGET) break;
      } catch { break; }
    }
    return best;
  };
  const applyEventPhotoToLogo = async (setter: (b64: string) => void): Promise<string> => {
    return await applyEventPhotoToLogoImpl({
      editEvent, fileToBase64, imageFile, imageOrigFile, imagePreview, isDe,
      showAlert, shrinkLogoB64,
    }, setter);
  };
  // v27.2: Größensteuerung fürs Kopfbild als wiederverwendbarer Block (Schritt 23
  // UND 24) — inkl. verkleinerter Live-Vorschau, die zeigt, wie groß das Bild im
  // Mail-/Outlook-Kopf steht. `headerImageLayout` gilt event-weit (Mail + Outlook).
  const renderHeaderSizeControl = (previewSrc: string, note?: string): React.ReactElement => {
    return renderHeaderSizeControlImpl({
      headerImageLayout, isDe, setHeaderImageLayout,
    }, previewSrc, note);
  };
  const [dragFieldId, setDragFieldId] = React.useState<string | null>(null);
  // v18.55: Pro-Feld Ein-/Ausklapp-Status für Schritt 5 (Felder). Default =
  // eingeklappt (kompakte Karte: nur Nummer + Label + Typ + Pflicht + Aktionen);
  // Detail-Einstellungen (Hilfetext, Optionen, Bedingung, CC, EN-Variante …)
  // erst beim Aufklappen. Neu hinzugefügte Felder starten aufgeklappt.
  const [fieldExpandOverride, setFieldExpandOverride] = React.useState<Record<string, boolean>>({});
  const toggleFieldExpand = (id: string, current: boolean): void =>
    setFieldExpandOverride(prev => ({ ...prev, [id]: !current }));
  const [dragOverFieldId, setDragOverFieldId] = React.useState<string | null>(null);
  // v9.28: Reorder-Mode toggelt die Hoch/Runter-Pfeile pro Custom-Field.
  // Standardmäßig aus — sonst sieht das Feld-Listing zu unruhig aus.
  const [reorderMode, setReorderMode] = React.useState(false);
  // v9.28: Modal für neuen Quiz-Bereich (statt window.prompt)
  const [newSectionModalOpen, setNewSectionModalOpen] = React.useState(false);
  const [newSectionName, setNewSectionName] = React.useState('');
  const [newSectionError, setNewSectionError] = React.useState('');
  const [agenda, setAgenda] = React.useState<AgendaItem[]>(
    editEvent && editEvent.agenda ? [...editEvent.agenda] : []
  );
  const [transferTimes, setTransferTimes] = React.useState<Array<{id: string; location: string; meetingPoint: string; address: string; date: string; departureTime: string; arrivalTime: string; description: string}>>(
    editEvent?.transferTimes?.map(t => ({...t, meetingPoint: t.meetingPoint || '', address: t.address || '', arrivalTime: t.arrivalTime || '', description: t.description || ''})) || []
  );
  const [documents, setDocuments] = React.useState<Array<{name: string; file?: File; url: string; size: number}>>(
    editEvent?.documents?.map(d => ({...d, size: d.size || 0})) || []
  );
  // Snapshot der beim Edit-Start vorhandenen Dokument-Namen, um beim Speichern
  // entfernte Attachments aus SharePoint löschen zu können.
  const [initialDocumentNames] = React.useState<string[]>(
    editEvent?.documents?.map(d => d.name) || []
  );
  const [quiz, setQuiz] = React.useState<Array<{id: string; question: string; options: string[]; correctIndices: number[]; imageBase64?: string; section?: string}>>(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    editEvent?.quiz?.map(q => ({...q, correctIndices: q.correctIndices || [(q as any).correctIndex || 0], imageBase64: (q as any).imageBase64, section: (q as any).section})) || []
  );
  const quizClusterSize = editEvent?.quizClusterSize || 1;
  // Sub-Events-Drafts im UI. Seit v6.4: Sub-Events sind eigene DEX_Events-Items.
  // Beim Edit laden wir die bestehenden Child-Events und mappen sie auf Drafts.
  // Beim Save werden Drafts mit `dbId` als updateEvent, ohne als createEvent geschrieben;
  // in der DB verbliebene Child-Events, die nicht mehr im Draft sind, werden gelöscht.
  const [subEvents, setSubEvents] = React.useState<SubEventDraft[]>(() => {
    if (!editEvent) return [];
    const kids = childEventsOf(editEvent.id);
    return kids.map(k => {
      // v11.57: Pro-Sub-Event Logo-Bilder aus EmailTemplateOverrides
      // (Piggyback-Pattern, gleich wie Top-Level-Event).
      // v14.4: zusätzlich die Mail-Text-Overrides (Anmeldung/Warteliste/
      // Abmeldung/Nachrücken) — vorher landeten Edits auf Sub-Event-Tabs
      // versehentlich beim Haupt-Event.
      let emailLogo = '';
      let outlookLogo = '';
      let subOverrides: Record<string, EmailOverrideEntry> = {};
      // v15.0: Inheritance-Flags aus dem Piggyback-JSON lesen. Wenn der
      // Flag nicht persistiert wurde (alte Events) fällt die App auf
      // datenbasierte Heuristik zurück (siehe weiter unten).
      let inheritFlagsRaw: { capacity?: boolean; fields?: boolean; location?: boolean } | undefined;
      try {
        const ov = JSON.parse(k.emailTemplateOverrides || '{}') as Record<string, unknown>;
        emailLogo = (ov?._eventLogo as string) || '';
        outlookLogo = readOutlookLogo(ov);
        inheritFlagsRaw = (ov?._inheritFlags as { capacity?: boolean; fields?: boolean; location?: boolean } | undefined);
        // Piggyback-Keys (mit Unterstrich-Prefix) rausstrippen, der Rest sind
        // die echten Mail-Template-Overrides pro TemplateType.
        const filtered: Record<string, EmailOverrideEntry> = {};
        for (const key of Object.keys(ov)) {
          if (key.startsWith('_')) continue;
          const val = ov[key] as Partial<EmailOverrideEntry> | undefined;
          if (val && (val.subject || val.heading || val.bodyHtml || val.headingColor || val.headingFontSize || val.subheading !== undefined || val.headingBold !== undefined || val.headingItalic !== undefined || val.subheadingColor || val.subheadingFontSize || val.subheadingBold !== undefined || val.subheadingItalic !== undefined)) {
            filtered[key] = {
              subject: val.subject || '',
              heading: val.heading || '',
              bodyHtml: val.bodyHtml || '',
              // v18.19/v18.22: Überschrift-Farbe/-Größe/-Stil + Subheading-
              // Formatierung mit-übernehmen.
              ...(val.subheading !== undefined ? { subheading: val.subheading } : {}),
              ...(val.headingColor ? { headingColor: val.headingColor } : {}),
              ...(val.headingFontSize ? { headingFontSize: val.headingFontSize } : {}),
              ...(val.headingBold !== undefined ? { headingBold: val.headingBold } : {}),
              ...(val.headingItalic !== undefined ? { headingItalic: val.headingItalic } : {}),
              ...(val.subheadingColor ? { subheadingColor: val.subheadingColor } : {}),
              ...(val.subheadingFontSize ? { subheadingFontSize: val.subheadingFontSize } : {}),
              ...(val.subheadingBold !== undefined ? { subheadingBold: val.subheadingBold } : {}),
              ...(val.subheadingItalic !== undefined ? { subheadingItalic: val.subheadingItalic } : {}),
            };
          }
        }
        subOverrides = filtered;
      } catch { /* */ }
      const parsedHeads = parseOutlookHeadings(k.outlookBody || '');
      // v15.0: Inheritance-Heuristik für Bestands-Events: wenn das
      // Piggyback-Flag fehlt UND das jeweilige Datenfeld nicht-leer ist,
      // gilt es als „eigener Wert" (nicht vom Hauptevent geerbt). Wenn
      // das Feld leer ist, default = übernehmen.
      const inheritCap = inheritFlagsRaw && typeof inheritFlagsRaw.capacity === 'boolean'
        ? inheritFlagsRaw.capacity
        : !(k.maxParticipants && k.maxParticipants > 0);
      const inheritFields = inheritFlagsRaw && typeof inheritFlagsRaw.fields === 'boolean'
        ? inheritFlagsRaw.fields
        : !((k.eventSpecificFields || []).length > 0);
      const inheritLoc = inheritFlagsRaw && typeof inheritFlagsRaw.location === 'boolean'
        ? inheritFlagsRaw.location
        : !(k.location && k.location.trim().length > 0);
      return {
      id: k.id,
      dbId: k.id,
      title: k.title,
      description: k.description,
      location: k.location,
      startDate: k.startDate,
      endDate: k.endDate,
      maxParticipants: k.maxParticipants || 0,
      registrationDeadline: k.registrationDeadline,
      mandatory: !!k.mandatoryRegistration, // v24.64: Pflicht-Sub-Event
      disableEmails: k.disableEmails,
      disableRegistrationEmail: k.disableRegistrationEmail,
      disableCancellationEmail: k.disableCancellationEmail,
      autoDeregisterOnDecline: k.autoDeregisterOnDecline,
      inactiveHandling: k.inactiveHandling,
      disableOutlook: k.disableOutlook,
      // v11.57: pro-Sub-Event Kommunikations-Felder laden
      emailLanguage: k.emailLanguage || (locale === 'de' ? 'DE' : 'EN'),
      emailLogoBase64: emailLogo,
      outlookLogoBase64: outlookLogo,
      emailTemplateOverrides: subOverrides,
      outlookBody: stripOutlookWrapper(k.outlookBody || ''),
      outlookHeading: parsedHeads.heading || k.title || '',
      outlookSubheading: parsedHeads.subheading && parsedHeads.subheading !== 'Event Details' ? parsedHeads.subheading : '',
      outlookSubject: k.outlookSubject || '',
      outlookStart: k.outlookStart || '',
      outlookEnd: k.outlookEnd || '',
      outlookLocation: k.outlookLocation || '',
      allDay: !!k.allDay,
      showAsFree: !!k.showAsFree,
      // v11.57: Snapshot der initialen Outlook-relevanten Felder
      initialOutlookEventId: k.outlookEventId || '',
      // v11.61: CalendarLink (iCalUId) als Outlook-Existenz-Indikator. Der
      // Flow schreibt OutlookEventId nicht — auf erfolgreichen Sub-Events
      // ist nur CalendarLink gefüllt.
      initialCalendarLink: k.calendarLink || '',
      initialTitle: k.title || '',
      initialAllDay: !!k.allDay,
      initialShowAsFree: !!k.showAsFree,
      initialStartDate: k.startDate || '',
      initialEndDate: k.endDate || '',
      initialOutlookBody: k.outlookBody || '',
      // v28.30: Kopfbild-Snapshot (gleicher Piggyback-Key wie beim Hauptevent).
      initialOutlookLogoBase64: ((): string => {
        if (!k.emailTemplateOverrides) return '';
        try { return readOutlookLogo(JSON.parse(k.emailTemplateOverrides)); } catch { return ''; }
      })(),
      customFields: (k.eventSpecificFields || []).map(f => ({
        id: f.id,
        label: f.label,
        type: f.type as CustomFieldInput['type'],
        required: !!f.required,
        options: f.options || [],
        // EventSpecificField hat kein 'visible'-Feld — default auf true.
        // Sichtbarkeit ist im Storage immer „shown" (default), nur im Wizard
        // kann der User Felder ausblenden.
        visible: true,
        externalLinks: f.externalLinks,
        multi: f.multi,
        helpText: f.helpText,
        helpTextStyle: f.helpTextStyle,
        showIf: f.showIf,
        // v29.20 (Audit A3): Dieser Mapper übernahm nur eine Teilmenge der
        // Feld-Eigenschaften — der nächste Save eines Klammer-Events schrieb
        // die Sub-Event-CustomFields dann OHNE den Rest zurück (die
        // v11.21-Drop-Klasse, hier im Lade-Pfad). „Felder vom Hauptevent
        // kopieren" und der Sub-Feld-Editor setzen all diese Properties.
        ...(f.onlyForGroup ? { onlyForGroup: f.onlyForGroup } : {}),
        ...(f.confirmLabel ? { confirmLabel: f.confirmLabel } : {}),
        ...(f.defaultValue ? { defaultValue: f.defaultValue } : {}),
        ...(f.optionCategories && f.optionCategories.length > 0 ? { optionCategories: [...f.optionCategories] } : {}),
        ...(f.prefilterLabel ? { prefilterLabel: f.prefilterLabel } : {}),
        ...(f.labelEn ? { labelEn: f.labelEn } : {}),
        ...(f.helpTextEn ? { helpTextEn: f.helpTextEn } : {}),
        ...(f.confirmLabelEn ? { confirmLabelEn: f.confirmLabelEn } : {}),
        ...(f.optionsEn && f.optionsEn.length > 0 ? { optionsEn: [...f.optionsEn] } : {}),
        ...(f.ccOnEmails ? { ccOnEmails: true } : {}),
        ...(f.notifyRoommate === false ? { notifyRoommate: false } : {}),
        ...(f.audienceOnly ? { audienceOnly: true } : {}),
        ...(f.withTime ? { withTime: true } : {}),
        ...(f.rangeStart ? { rangeStart: f.rangeStart } : {}),
        ...(f.rangeEnd ? { rangeEnd: f.rangeEnd } : {}),
        ...(typeof f.maxNights === 'number' && f.maxNights > 0 ? { maxNights: f.maxNights } : {}),
      })),
      // v15.3: pro-Sub-Event Felder aus dem Event-Datenmodell laden. Alle
      // Sub-Events haben jetzt eigene Adresse, Agenda, Transferzeiten,
      // Deadline, Standortfilter, Audience, Filter-Modus, Warteliste und
      // Anrede-Toggle — wie der Hauptevent.
      locationAddress: k.locationAddress ? {
        street: k.locationAddress.street || '',
        houseNo: k.locationAddress.houseNo || '',
        zip: k.locationAddress.zip || '',
        city: k.locationAddress.city || '',
      } : { street: '', houseNo: '', zip: '', city: '' },
      agenda: (k.agenda || []) as AgendaItem[],
      transferTimes: (k.transferTimes || []).map(tt => ({
        id: tt.id,
        location: tt.location || '',
        meetingPoint: tt.meetingPoint || '',
        address: tt.address || '',
        date: tt.date || '',
        departureTime: tt.departureTime || '',
        arrivalTime: tt.arrivalTime || '',
        description: tt.description || '',
      })),
      lastDeregisterDate: k.lastDeregisterDate || '',
      // Form-Felder für Standortfilter / Mailverteiler sind comma-separated
      // Strings, persistiert im Event aber als Arrays — siehe Top-Level-Mapping.
      locationFilter: (k.locationAudience || []).join(', '),
      audience: (k.audienceFilter || []).join(', '),
      filterMode: (k.filterMode === 'AND' ? 'AND' : 'OR') as 'AND' | 'OR',
      // v22.10: Ausschluss-Liste des Sub-Events laden (vorher nicht persistiert).
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      excludedUsers: ((k as any).excludedUsers || []) as string[],
      waitlistEnabled: typeof k.waitlistEnabled === 'boolean' ? k.waitlistEnabled : true,
      askSalutation: !!k.askSalutation,
      // v27.11: bestehendes Sub-Event-Bild (SP-URL) als Vorschau laden.
      imagePreview: k.imageUrl || '',
      // v15.0 (legacy): Inheritance-Flags werden seit v15.3 nicht mehr
      // ausgewertet. Bleiben in den geparsten Drafts, weil das Schema
      // sie noch erlaubt — Wirkung gleich Null.
      inheritLocationFromParent: inheritLoc,
      inheritCapacityFromParent: inheritCap,
      inheritCustomFieldsFromParent: inheritFields,
    };
    });
  });
  // v29.75: Solange der „Sichtbarkeit gilt für alle Sub-Events"-Haken gesetzt
  // ist, spiegelt jede Änderung an Standortfilter/Verteiler/Verknüpfung der
  // Klammer sofort in alle Sub-Event-Drafts — auch das Setzen des Hakens
  // selbst. Die Sub-Event-Sichtbarkeits-UI ist währenddessen gesperrt (sonst
  // würden dortige Eingaben beim nächsten Klammer-Edit stillschweigend
  // überschrieben). Unveränderte Drafts behalten ihre Referenz, damit der
  // v29.57-Skip („nichts geändert → nicht schreiben") nicht anschlägt.

  return {
    agenda, applyEventPhotoToLogo, assistantsCanSee, documents, dragFieldId, dragOverFieldId,
    effTeamsLink, emailLogoPreview, emailTemplateOverrides, emailTemplates, fieldExpandOverride, fileToBase64,
    headerImageLayout, headerImageLayoutConfig, headerLayoutFor, hiddenOrganizerEmails, hideOrganizer, hideOrganizerIndividualOnly,
    initialDocumentNames, initialHeaderImageLayoutRef, isFictive, newSectionError, newSectionModalOpen, newSectionName,
    onlineMeetingChanged, organizerDisplayLarge, outlookLogoPreview, outlookTeamsLink, quiz, quizClusterSize,
    renderHeaderSizeControl, reorderMode, setAgenda, setAssistantsCanSee, setDocuments, setDragFieldId,
    setDragOverFieldId, setEmailLogoPreview, setEmailTemplateOverrides, setEmailTemplates, setFieldExpandOverride, setHeaderImageLayout,
    setHideOrganizer, setHideOrganizerIndividualOnly, setIsFictive, setNewSectionError, setNewSectionModalOpen, setNewSectionName,
    setOrganizerDisplayLarge, setOutlookLogoPreview, setQuiz, setReorderMode, setSubEvents, setTransferTimes,
    setTriggerOutlookUpdate, shrinkLogoB64, subEvents, toggleFieldExpand, toggleOrganizerHidden, transferTimes,
    triggerOutlookUpdate,
  };
}
