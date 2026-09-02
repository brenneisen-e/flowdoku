/* persistSubEvents — aus EventCreationPage.tsx ausgelagert (Zeilen 3450-4011 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
import * as React from 'react';
import { formatOrganizerList } from '../../../context/EventContext';
import { buildOutlookBody, getCachedOrbBase64, normalizeMadeWithLink, replacePlaceholders } from '../../../services/EmailTemplates';
import { outlookLogoPiggyback, serializeCustomFields } from '../../wizard/wizardHelpers';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { isThrottled } from '../../../utils/spThrottle';
import { dlog } from '../../../utils/debugLog';
import { SubEventDraft } from '../../wizard/wizardTypes';
import { EmailOverrideEntry } from '../../wizard/emailOverrideEntry';

export interface PersistSubEventsCtx {
  // v30.67: Adresse des Hauptevents — Fallback für {{Address}} im Outlook-
  // Body eines Sub-Events ohne eigene Adresse.
  addrCity: string;
  addrHouseNo: string;
  addrStreet: string;
  addrZip: string;
  bilingualFields: boolean;
  childEventsOf: (parentEventId: string) => import("../../../types/index").DeloitteEvent[];
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  contactEmail: string;
  createEvent: (event: import("../../../context/eventContextTypes").CreateEventInput) => Promise<number>;
  deleteEvent: (eventId: string) => Promise<boolean>;
  deleteEventItemOnly: (eventId: string) => Promise<boolean>;
  editEvent: import("../../../types/index").DeloitteEvent;
  forceOutlookRecreateRef: React.MutableRefObject<Set<string>>;
  headerImageLayoutConfig: { _headerImageLayout: { width: number; paddingV: number; paddingH: number; }; } | { _headerImageLayout?: undefined; };
  headerLayoutFor: (logoB64: string) => {    imageWidth: number;    imagePaddingV: number;    imagePaddingH: number;};
  initialSubEventDbIds: string[];
  initialSubEventOutlookMeta: Record<string, { disableOutlook: boolean; outlookEventId: string; subsiteUrl: string; registrationListName: string; }>;
  initialSubPersistRef: React.MutableRefObject<Record<string, string>>;
  isDe: boolean;
  isFictive: boolean;
  onlineMeetingMode: "none" | "own" | "auto";
  organizer: string;
  orgGetsSubInvites: boolean;
  outlookTeamsLink: () => string;
  parentTimesIso: () => {    start: string;    end: string;};
  pendingOutlookRecreateForSubEventsRef: React.MutableRefObject<string[]>;
  persistSubEventImage: (subDbId: string | number | null | undefined, draft: { imageFile?: File | null; imageRemoved?: boolean; }) => Promise<void>;
  /** v30.67 (Review): einmaliger Reload nach einem Recreate (s. Funktionsende). */
  refreshEvents: () => Promise<void>;
  resolveTopLevelCommState: () => { emailLanguage: string; emailLogoBase64: string; outlookLogoBase64: string; outlookBody: string; outlookHeading: string; outlookSubheading: string; outlookSubject: string; disableEmails: boolean; disableRegistrationEmail: boolean; disableCancellationEmail: boolean; autoDeregisterOnDecline: boolean; inactiveHandling?: 'notify' | 'autoderegister'; disableOutlook: boolean; emailTemplateOverrides: Record<string, EmailOverrideEntry>; };
  sanitizeOrganizerPairs: () => {    orgString: string;    orgEmailString: string;    droppedCount: number;};
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  shrinkLogoB64: (b64: string) => Promise<string>;
  subEventCalendar: boolean;
  subEventsRef: React.MutableRefObject<SubEventDraft[]>;
  subPersistKey: (d: SubEventDraft) => string;
  subPhotoAsLogo: (draft: { imageFile?: File | null; imagePreview?: string; imageRemoved?: boolean; }) => Promise<string>;
  subTopGateInitialRef: React.MutableRefObject<string>;
  subTopGateKey: () => string;
  title: string;
  updateEvent: (eventId: string, updates: Record<string, unknown>, opts?: { skipReload?: boolean; }) => Promise<boolean>;
}

export async function persistSubEventsForParentImpl(ctx: PersistSubEventsCtx, parentEventId: string, onStep: (done: number, total: number, title: string) => void): Promise<void> {
  const { addrCity, addrHouseNo, addrStreet, addrZip, bilingualFields, childEventsOf, confirmDialog, contactEmail, createEvent, deleteEvent, deleteEventItemOnly, editEvent, forceOutlookRecreateRef, headerImageLayoutConfig, headerLayoutFor, initialSubEventDbIds, initialSubEventOutlookMeta, initialSubPersistRef, isDe, isFictive, onlineMeetingMode, organizer, orgGetsSubInvites, outlookTeamsLink, parentTimesIso, pendingOutlookRecreateForSubEventsRef, persistSubEventImage, refreshEvents, resolveTopLevelCommState, sanitizeOrganizerPairs, showAlert, shrinkLogoB64, subEventCalendar, subEventsRef, subPersistKey, subPhotoAsLogo, subTopGateInitialRef, subTopGateKey, title, updateEvent } = ctx;
    const keptDbIds = new Set<string>();
    // v30.67: Ids, deren DEX_Events-Zeile in DIESEM Lauf per Recreate ersetzt
    // (oder auf dem Legacy-Pfad bewusst gelöscht) wurde. Sie sind weder
    // „behalten" (die Zeile ist weg) noch „abgewählt" (der Termin lebt unter
    // neuer Id weiter) — die Aufräum-Schleife unten darf sie deshalb NIE
    // anfassen. Vorher rief sie für jede ersetzte Id deleteEvent (→ 404 →
    // false) und meldete nach einem vollständig erfolgreichen Speichern
    // „N Termine konnten nicht gelöscht werden", mit je einem kompletten
    // Event-Reload obendrauf.
    const replacedDbIds = new Set<string>();
    const failedSubTitles: string[] = [];
    // v29.57: Einmal je Save auswerten, nicht je Sub-Event — die Schranke
    // hängt nur am Hauptevent. Bei einem NEUEN Event (kein Snapshot) wird nie
    // übersprungen; dort gibt es ohnehin nichts zu vergleichen.
    const subGateUnchanged = !!editEvent
      && subTopGateInitialRef.current !== ''
      && subTopGateInitialRef.current === subTopGateKey();
    let skippedSubCount = 0;
    // v29.74: Drossel-Abbruch (s. Schleifenrumpf).
    let consecutiveSubFailures = 0;
    let abortedForThrottle = false;
    const stepTotal = subEventsRef.current.filter(d => !!(d.title || '').trim()).length;
    let stepDone = 0;
    // v11.87: Sub-Event-Progress-Callback aus dem aufrufenden handleSubmit
    // einspeisen. Der Caller setzt window.__dexSubEventProgress vor dem
    // Aufruf und entfernt es danach. Wenn nicht gesetzt: no-op.
    const subOnProgress = (stage: string): void => {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const cb = (window as any).__dexSubEventProgress;
        if (typeof cb === 'function') cb(stage);
      } catch { /* */ }
    };
    // v30.67 (Review): Recreate mit Subsite-Reuse — ERST die neue Zeile
    // anlegen, DANN die alte löschen. Vorher lief es umgekehrt: Scheiterte
    // createEvent nach dem Delete (429 auf dem ersten GET, 2-MB-Grenze, POST
    // abgelehnt), war die DEX_Events-Zeile des Termins weg — Subsite verwaist,
    // Termin in keiner Ansicht — und der Wizard meldete „gespeichert", weil
    // der catch nur in die Konsole schrieb. Jetzt bleibt bei JEDEM Fehlschlag
    // die alte Zeile stehen; der Aufrufer fällt in den normalen Update-Pfad
    // (dort landet die dbId in keptDbIds) und der Recreate-Merker bleibt für
    // den nächsten Save erhalten. Schlägt erst das Löschen der ALTEN Zeile
    // fehl, wird die neue zurückgenommen; gelingt auch das nicht, stehen zwei
    // Zeilen auf derselben Subsite — die alte wird behalten, der Organizer
    // bekommt den Hinweis, den doppelten Eintrag NICHT selbst zu löschen
    // (deleteEvent würde die geteilte Subsite mitnehmen).
    // Rückgabe true = ersetzt (Aufrufer macht `continue`), false = Update-Pfad.
    const recreateWithReuse = async (draft: SubEventDraft, reusePayload: unknown, logTag: string): Promise<boolean> => {
      let recreatedId: number | null = null;
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        recreatedId = await createEvent(reusePayload as any);
      } catch (err) {
        console.warn(`[DEX][${logTag}] Sub-Event-Recreate mit Subsite-Reuse fehlgeschlagen — alte Zeile bleibt:`, draft.dbId, err);
        recreatedId = null;
      }
      if (!recreatedId) {
        showAlert(isDe
          ? `Der Outlook-Termin für „${draft.title || 'Sub-Event'}" konnte nicht neu angelegt werden (SharePoint hat den neuen Eintrag abgelehnt). Das Sub-Event und seine Anmeldungen sind unverändert — bitte versuche es später erneut.`
          : `The Outlook appointment for "${draft.title || 'sub-event'}" could not be recreated (SharePoint rejected the new entry). The sub-event and its registrations are untouched — please try again later.`, { variant: 'error' });
        return false;
      }
      const oldGone = await deleteEventItemOnly(draft.dbId).catch(() => false);
      if (oldGone) {
        // Die alte Zeile ist nachweislich weg — nie mehr in den kaskadierenden
        // deleteEvent-Zweig (s. replacedDbIds).
        replacedDbIds.add(draft.dbId);
        // v27.11: Sub-Event-Bild auch auf dem Recreate-Pfad persistieren.
        await persistSubEventImage(recreatedId, draft);
        return true;
      }
      const rolledBack = await deleteEventItemOnly(String(recreatedId)).catch(() => false);
      console.warn(`[DEX][${logTag}] Recreate: alte Zeile nicht löschbar, neue Zeile ${rolledBack ? 'zurückgenommen' : 'NICHT zurückgenommen'}:`, draft.dbId, recreatedId);
      showAlert(rolledBack
        ? (isDe
          ? `Der Outlook-Termin für „${draft.title || 'Sub-Event'}" konnte nicht neu angelegt werden (SharePoint hat den Austausch des Eintrags abgelehnt). Das Sub-Event und seine Anmeldungen sind unverändert — bitte versuche es später erneut. Möglicherweise liegt trotzdem schon ein Outlook-Termin für den zurückgenommenen Eintrag in deinem Kalender — bitte lösche ihn dort von Hand.`
          : `The Outlook appointment for "${draft.title || 'sub-event'}" could not be recreated (SharePoint rejected replacing the entry). The sub-event and its registrations are untouched — please try again later. An Outlook appointment for the withdrawn entry may nevertheless already be in your calendar — please delete it there manually.`)
        : (isDe
          ? `Beim Neuanlegen des Outlook-Termins für „${draft.title || 'Sub-Event'}" ist ein doppelter Eintrag entstanden (Eintrag-Id ${recreatedId}). Die Anmeldungen sind unverändert. Bitte melde dich beim DEX-Team, damit der doppelte Eintrag entfernt wird — lösche ihn NICHT selbst, weder im Organizer Center noch über „Entfernen" im Wizard: Das würde die Anmeldungen mitlöschen.`
          : `Recreating the Outlook appointment for "${draft.title || 'sub-event'}" left a duplicate entry (item id ${recreatedId}). Registrations are unchanged. Please contact the DEX team to remove the duplicate — do NOT delete it yourself, neither in the Organizer Center nor via "Remove" in the wizard: that would delete the registrations as well.`),
        { variant: 'error' });
      return false;
    };
    // Sub-Events erben Organizer + OrganizerEmail vom Parent. Einmal sanitisieren
    // statt pro Iteration, identisch für alle Children.
    const sanitizedOrgPair = sanitizeOrganizerPairs();
    // v28.66 BUG-FIX: Zeiten des Hauptevents als Fallback für Sub-Events ohne
    // eigene Zeiten. Die Start-/End-Felder eines Sub-Events sind optional (neue
    // Drafts starten auf '' und der DatePicker ist löschbar) und werden — anders
    // als beim Hauptevent — von der Wizard-Prüfung nicht eingefordert. Bisher
    // landeten StartDate UND EndDate dann leer in DEX_Events; der
    // DEX_CreateOutlookEvent-Flow rechnet convertFromUtc(coalesce(OutlookEnd,
    // EndDate)) = convertFromUtc(null) und „Create event (V4)" bricht ab — für
    // dieses Sub-Event entsteht gar kein Outlook-Termin (und der Flow-Lauf
    // scheitert komplett). Ein Sub-Event findet immer innerhalb seines
    // Hauptevents statt, also sind dessen Zeiten der richtige Default. Nach dem
    // Speichern stehen sie sichtbar in den Sub-Event-Feldern und können dort
    // präzisiert werden.
    const { start: parentStartIso, end: parentEndIso } = parentTimesIso();
    // v28.29 BUG-FIX: Kopfbild-Vererbung vom Hauptevent auf die Sub-Events.
    // Sub-Events haben EIGENE Outlook-Termine und eigene Mails, aber praktisch
    // nie ein eigenes Kopfbild — Schritt 23/24 wird pro Tab gepflegt, und die
    // Vorschau dort zeigte fälschlich das Event-Foto, obwohl im Sub-Tab gar
    // nichts hinterlegt war. Beim Speichern fiel der Sub-Event-Body deshalb
    // auf das Standard-DEX-Logo zurück: Hauptevent-Termin mit Foto,
    // Sub-Event-Termine mit Orb. Jetzt erbt jedes Sub-Event ohne eigenes Bild
    // das Bild des Hauptevents; ein eigenes Bild im Sub-Tab gewinnt weiterhin.
    const parentComm = resolveTopLevelCommState();
    const inheritedEmailLogo = await shrinkLogoB64(parentComm.emailLogoBase64 || '');
    const inheritedOutlookLogo = await shrinkLogoB64(parentComm.outlookLogoBase64 || '');
    // v11.60: aus dem Ref iterieren — der React-State ist beim Save evtl.
    // noch nicht propagiert, weil flushActiveCommTabToState() per setState
    // erst async wirkt. Der Ref hält synchron die letzten Tab-Werte.
    for (const draft of subEventsRef.current) {
      if (!draft.title || !draft.title.trim()) {
        // v29.19: Leere Drafts werden nicht gespeichert — aber ein BESTEHENDES
        // Sub-Event (dbId), dessen Titel gerade nur geleert ist, gilt trotzdem
        // als „behalten". Vorher fiel es in die Lücke zwischen zwei
        // „behalten"-Definitionen: Der Warn-Dialog in handleSubmitInner zählt
        // über dbIds (Draft existiert → kein Dialog), die Aufräum-Schleife
        // unten über keptDbIds (nicht drin → deleteEvent) — das Sub-Event
        // wurde samt Subsite und Anmeldungen OHNE Rückfrage gelöscht, nur
        // weil jemand den Titel zum Neutippen geleert und dann gespeichert
        // hat. Löschen geht weiterhin — über das X an der Karte, mit Dialog.
        if (draft.dbId) keptDbIds.add(draft.dbId);
        continue;
      }
      // v11.57: Pro-Sub-Event Kommunikations-Felder. Wenn der Organizer für
      // den Sub-Event eigene Werte in Step 5 gesetzt hat, verwenden wir die;
      // sonst fallback auf die Top-Level-Werte (Backward-Compat für
      // Sub-Events ohne eigene Communication-Einstellungen).
      // v30.67: Fallback über `parentComm`, nicht über den rohen State — auf
      // einem Sub-Reiter hält `emailLanguage` den Wert DIESES Reiters, und
      // ein neuer Termin (ohne eigene Sprache) erbte beim Speichern von dort
      // die Sprache des zuletzt geöffneten Termins statt die des Hauptevents.
      const subEmailLang = draft.emailLanguage || parentComm.emailLanguage;
      const subOutlookBodyRaw = (typeof draft.outlookBody === 'string' && draft.outlookBody !== '') ? draft.outlookBody : '';
      const subOutlookHeading = draft.outlookHeading || draft.title || '';
      const subOutlookSub = draft.outlookSubheading || '';
      // v30.7: Kalender-Tage heissen „Di. 01.09.2026" — ohne eigenen Betreff
      // fiel der Outlook-Flow auf diesen Tages-Titel zurueck, und im Kalender
      // der Teilnehmer stand nur das Datum (das der Termin ohnehin traegt).
      // Default ist deshalb der NAME DES HAUPTEVENTS; ein im Kommunikations-
      // Reiter gesetzter eigener Betreff gewinnt weiterhin.
      const subOutlookSubject = (draft.outlookSubject || '').trim() || (subEventCalendar ? title.trim() : '');
      // v28.29: eigenes Bild des Sub-Events gewinnt, sonst erbt es das
      // Kopfbild des Hauptevents (statt still auf den Orb zu fallen).
      // v29.20 (Audit): auch das EIGENE Sub-Logo verkleinern — der
      // v28.10-Schutz lief nur über die geerbten Parent-Logos. Ein auf dem
      // Sub-Reiter hochgeladenes unkomprimiertes Foto steckte bis zu dreimal
      // im Payload und riss das Sub-Event ins SharePoint-2-MB-Limit.
      // v29.32: Zwischen eigenem Logo und geerbtem Parent-Logo steht jetzt das
      // EIGENE Bild des Sub-Events (s. subPhotoAsLogo) — bei Terminreihen mit
      // Foto je Termin kam sonst überall das Bild des Hauptevents an. Nur
      // aufgerufen, wenn kein eigenes Logo gesetzt ist (spart Laden/Komprimieren).
      const subOwnPhotoLogo = draft.emailLogoBase64 ? '' : await subPhotoAsLogo(draft);
      const subEmailLogo = (draft.emailLogoBase64 ? await shrinkLogoB64(draft.emailLogoBase64) : '') || subOwnPhotoLogo || inheritedEmailLogo;
      const subOutlookLogo = (draft.outlookLogoBase64 ? await shrinkLogoB64(draft.outlookLogoBase64) : '') || subOwnPhotoLogo || inheritedOutlookLogo;
      // Outlook-Body wrappen. v26.59 BUG-FIX: Ohne eigenen Text wurde der Body
      // bisher LEER gespeichert („der Flow setzt einen Default" — stimmte
      // nicht, der Flow mappt 1:1) → die Outlook-Einladung der Sub-Events kam
      // komplett ohne Text an. Jetzt bekommen Sub-Events denselben
      // Default-Body wie das Haupt-Event (v7.4-Pattern: Anmeldebestätigung +
      // Abmelde-Hinweis über die App + Organizer-Kontakt).
      // v15.3: Sub-Event-eigene strukturierte Adresse serialisieren (analog
      // zum Hauptevent-Top-Level-Pattern). Wenn alle vier Komponenten leer
      // sind, wird ein leerer String gespeichert.
      // v30.67: nach oben gezogen — {{Address}} im Outlook-Body braucht sie
      // schon beim Aufbau der `vars` (stand dort hart auf '').
      const draftAddr = draft.locationAddress || { street: '', houseNo: '', zip: '', city: '' };
      const draftHasAddress = !!(draftAddr.street || draftAddr.houseNo || draftAddr.zip || draftAddr.city);
      let wrappedSubOutlookBody = '';
      {
        const orgNamesSub = formatOrganizerList([organizer], subEmailLang);
        // v30.67: {{Address}} war im Sub-Event-Zweig immer leer — der
        // Variablen-Knopf im Editor bot sie an, die Vorschau zeigte sie, im
        // Termin der Teilnehmer stand „Treffpunkt: “. Dieselbe Formatierung
        // wie im Hauptevent-Pfad; ohne eigene Adresse die des Hauptevents
        // (ein Sub-Event findet innerhalb seines Hauptevents statt).
        const fmtAddr = (street: string, houseNo: string, zip: string, city: string): string =>
          [street, houseNo].filter(Boolean).join(' ') + ((zip || city) ? ', ' + [zip, city].filter(Boolean).join(' ') : '');
        const subAddress = draftHasAddress
          ? fmtAddr(draftAddr.street || '', draftAddr.houseNo || '', draftAddr.zip || '', draftAddr.city || '')
          : fmtAddr(addrStreet, addrHouseNo, addrZip, addrCity);
        const vars = {
          EventTitle: draft.title.trim(),
          // v27.5: {{Organizer}} auf normalisierte Namen ("Vorname Nachname",
          // mit „und" verbunden) statt roher „Nachname, Vorname"-Join.
          Organizer: orgNamesSub || organizer,
          ContactEmail: contactEmail.trim(),
          Location: draft.location || '',
          Address: subAddress,
          StartDate: draft.startDate ? new Date(draft.startDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
          EndDate: draft.endDate ? new Date(draft.endDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
        };
        const escHtmlSub = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
        const APP_URL_SUB = 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView';
        const defaultSubBody = subEmailLang === 'EN'
          ? `<p>You are registered for the event <strong>${escHtmlSub(draft.title.trim())}</strong>.</p>`
            + `<p>If you are unable to attend, please cancel your registration in time via the <a href="${APP_URL_SUB}" style="color:#86bc25;font-weight:600;">DEX App</a> (&bdquo;My Events&ldquo;).</p>`
            + `<p>For organizational questions please contact <strong>${escHtmlSub(orgNamesSub || 'the organizer')}</strong>.</p>`
          : `<p>Ihr seid für das Event <strong>${escHtmlSub(draft.title.trim())}</strong> angemeldet.</p>`
            + `<p>Falls ihr nicht teilnehmen könnt, meldet euch bitte rechtzeitig über die <a href="${APP_URL_SUB}" style="color:#86bc25;font-weight:600;">DEX App</a> (&bdquo;Meine Events&ldquo;) ab.</p>`
            + `<p>Bei organisatorischen Fragen wendet euch bitte an <strong>${escHtmlSub(orgNamesSub || 'den Organizer')}</strong>.</p>`;
        const resolvedBody = subOutlookBodyRaw ? replacePlaceholders(subOutlookBodyRaw, vars) : defaultSubBody;
        const resolvedHead = subOutlookHeading ? replacePlaceholders(subOutlookHeading, vars) : draft.title.trim();
        // v27.5: Default-Unter-Überschrift = Ort (nicht Datum).
        const resolvedSub2 = subOutlookSub ? replacePlaceholders(subOutlookSub, vars) : (draft.location || undefined);
        // v18.73: Sub-Events erben das Header-Bild-Layout des Hauptevents.
        // v28.29: ohne eigenes/geerbtes Bild wird die Breite gekappt (Orb).
        const wrapped = buildOutlookBody(resolvedHead, resolvedBody, resolvedSub2, headerLayoutFor(subOutlookLogo), outlookTeamsLink(), (subEmailLang || '').toUpperCase() !== 'EN');
        wrappedSubOutlookBody = wrapped.replace(/\{\{ORB_URL\}\}/g, subOutlookLogo || getCachedOrbBase64() || '');
      }
      // Sub-Event-EmailTemplateOverrides: Logo-Piggybacks (Top-Level-Pattern)
      // + ab v14.4 die echten Mail-Text-Overrides pro Sub-Event
      // (Anmeldung/Warteliste/Abmeldung/Nachrücken).
      const subDraftOverrides = draft.emailTemplateOverrides || {};
      // v15.3: Inheritance-Flags entfallen — Sub-Events sind seit v15.3
      // vollwertige Events mit eigener Konfiguration. Der Piggyback-Key
      // `_inheritFlags` wird nicht mehr geschrieben.
      const subOverridesMerged: Record<string, unknown> = {
        ...subDraftOverrides,
        ...(subEmailLogo ? { _eventLogo: subEmailLogo } : {}),
        ...outlookLogoPiggyback(subEmailLogo, subOutlookLogo),
        // v18.73: Sub-Event erbt das Header-Bild-Layout des Hauptevents, damit
        // auch die Sub-Event-Mails den gleichen Bild-Kopf nutzen.
        ...headerImageLayoutConfig,
      };
      const subEmailOverrides = Object.keys(subOverridesMerged).length > 0
        ? JSON.stringify(subOverridesMerged)
        : '';
      const draftLocationAddress = draftHasAddress ? JSON.stringify(draftAddr) : '';
      const draftAgendaJson = JSON.stringify(draft.agenda || []);
      const draftTransfersJson = JSON.stringify(draft.transferTimes || []);
      const childPayload = {
        title: draft.title.trim(),
        type: 'Other',
        status: 'Active',
        description: draft.description || '',
        location: draft.location || '',
        locationAddress: draftLocationAddress,
        locationFilter: draft.locationFilter || '',
        audience: draft.audience || '',
        filterMode: (draft.filterMode === 'AND' ? 'AND' : 'OR'),
        // v22.10: Ausschluss-Liste pro Sub-Event mitpersistieren (createEvent
        // schreibt sie in die Spalte ExcludedUsers).
        excludedUsers: draft.excludedUsers || [],
        // v28.66: ohne eigene Zeit die Zeit des Hauptevents erben (s.o.).
        startDate: draft.startDate || parentStartIso || '',
        // v22.17: NIE ein leeres EndDate persistieren — sonst rechnet der
        // DEX_CreateOutlookEvent-Flow convertFromUtc(coalesce(OutlookEnd,
        // EndDate)) = convertFromUtc(null) und „Create event (V4)" stürzt ab
        // (kein Outlook-Termin). Fallback auf das Start-Datum.
        // v28.66: greift auch, wenn das Sub-Event gar keine eigene Zeit hat —
        // dann kommt das Ende (ersatzweise der Start) des Hauptevents.
        endDate: draft.endDate || draft.startDate || parentEndIso || parentStartIso || '',
        registrationDeadline: draft.registrationDeadline || '',
        lastDeregisterDate: draft.lastDeregisterDate || '',
        maxParticipants: draft.maxParticipants || 0,
        waitlistEnabled: typeof draft.waitlistEnabled === 'boolean' ? draft.waitlistEnabled : true,
        mandatoryRegistration: !!draft.mandatory, // v24.64: Pflicht-Sub-Event
        eventImageUrl: '',
        organizer: sanitizedOrgPair.orgString,
        organizerEmail: sanitizedOrgPair.orgEmailString,
        outlookEventId: '',
        outlookBody: wrappedSubOutlookBody,
        outlookSubject: subOutlookSubject || undefined,
        outlookStart: (draft.outlookStart || '') || undefined,
        outlookEnd: (draft.outlookEnd || '') || undefined,
        outlookLocation: (draft.outlookLocation || '') || undefined,
        // v29.52: ganztägig mitschreiben — sonst kippt der Haken beim Speichern zurück.
        allDay: !!draft.allDay,
        showAsFree: !!draft.showAsFree,
        // v30.26: Die Online-Meeting-Entscheidung gilt event-weit — jeder
        // Termin einer Reihe bekommt dann seine eigene Teams-Besprechung.
        outlookIsOnlineMeeting: onlineMeetingMode === 'auto',
        skipOrganizerInvite: !orgGetsSubInvites, // v29.55
        agenda: draftAgendaJson,
        transfers: draftTransfersJson,
        documents: '[]',
        funZone: '[]',
        quizClusterSize: 1,
        emailLanguage: subEmailLang,
        emailTemplateOverrides: subEmailOverrides,
        disableEmails: !!draft.disableEmails,
        // v19.22: granulare An-/Abmelde-Mail-Schalter pro Sub-Event persistieren.
        disableRegistrationEmail: !!draft.disableRegistrationEmail,
        disableCancellationEmail: !!draft.disableCancellationEmail,
        autoDeregisterOnDecline: !!draft.autoDeregisterOnDecline,
        inactiveHandling: (draft.inactiveHandling === 'autoderegister' ? 'autoderegister' : 'notify') as 'notify' | 'autoderegister',
        disableOutlook: !!draft.disableOutlook,
        isFictive: isFictive,
        askSalutation: !!draft.askSalutation,
        // v29.20 (Audit): kanonisch serialisieren wie der Update-Zweig —
        // roh gingen leere Felder und leere Options-Slots (neue Drafts
        // starten mit ['','']) direkt auf die Anmeldeseite, EN-Varianten
        // wurden am bilingual-Schalter vorbei geschrieben und Wizard-interne
        // Properties (visible) landeten im Storage-JSON.
        customFields: serializeCustomFields(draft.customFields || [], bilingualFields),
        parentEventId: parentEventId,
      };
      if (draft.dbId) {
        // v11.69: Recreate-Pfad via Modal-Auswahl. Wenn der Organizer im
        // Outlook-Confirm-Modal ein Sub-Event mit `noOutlookYet=true`
        // angehakt hat (es existiert noch kein Outlook-Termin), wird das
        // bestehende DEX_Events-Item per `deleteEventItemOnly` entfernt und
        // eine NEUE DEX_Events-Zeile angelegt — wobei die bestehende
        // Subsite + Teilnehmerliste an die neue Zeile gekoppelt werden.
        // Damit triggert der `DEX_CreateOutlookEvent`-Flow (GetOnNewItems)
        // auf dem neuen Item und legt den Outlook-Termin an. Die alte
        // Subsite mit ALLEN Anmeldungen bleibt unangetastet.
        const initialMeta = initialSubEventOutlookMeta[draft.dbId];
        if (pendingOutlookRecreateForSubEventsRef.current.includes(draft.dbId)) {
          const subsiteUrlForReuse = initialMeta?.subsiteUrl || '';
          const regListNameForReuse = initialMeta?.registrationListName || 'Teilnehmer';
          if (subsiteUrlForReuse && regListNameForReuse) {
            // v29.21 (Audit): Den Rückgabewert PRÜFEN — deleteEventItemOnly
            // wirft nie, es meldet Fehlschlag als false (das catch hier war
            // toter Code). Schlug der Delete fehl (429/403), existierte die
            // alte Zeile weiter, createEvent legte eine ZWEITE Zeile auf
            // dieselbe Subsite, und die Aufräum-Schleife am Ende rief
            // deleteEvent auf die alte Id — KASKADIEREND: Sie recycelte die
            // geteilte Subsite mit allen Anmeldungen, auf die die frische
            // Zeile zeigt.
            // v30.67 (Review): Reihenfolge und Fehlerbehandlung liegen jetzt
            // in recreateWithReuse (erst anlegen, dann löschen). Bei false in
            // den normalen Update-Pfad fallen — die dbId landet dort in
            // keptDbIds, es geht nichts verloren.
            // Reuse-Payload: bestehende Subsite + Teilnehmerliste an die
            // neue DEX_Events-Zeile koppeln. disableOutlook explizit false,
            // outlookEventId leer, damit der Flow sauber neu schreibt.
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const reusePayload: any = {
              ...childPayload,
              disableOutlook: false,
              outlookEventId: '',
              existingSubsiteUrl: subsiteUrlForReuse,
              existingRegistrationListName: regListNameForReuse,
              onProgress: subOnProgress,
            };
            // NICHT zu keptDbIds hinzufügen, wenn ersetzt — das alte Item
            // wurde gelöscht und die neue Zeile hat eine andere ID.
            if (await recreateWithReuse(draft, reusePayload, 'v11.69')) continue;
          } else {
            console.warn('[DEX][v11.69] Recreate angefordert aber keine subsiteUrl/registrationListName vorhanden — Sub-Event:', draft.dbId, 'meta:', initialMeta);
            // Fall durch zum normalen Update-Pfad — wenigstens die Felder
            // werden persistiert; Outlook-Termin entsteht aber nicht.
          }
        }
        // Spezialfall: Outlook nachträglich aktivieren via DisableOutlook-
        // Toggle (alter Pfad vor v11.69). Wenn der User auf einem
        // **bestehenden** Sub-Event die "Outlook erstellen"-Checkbox
        // einschaltet (DisableOutlook: true → false) und bisher kein
        // Outlook-Termin angelegt wurde (OutlookEventId leer), muss das
        // Sub-Event neu angelegt werden — der Power-Automate-Flow
        // `DEX_CreateOutlookEvent` triggert ausschließlich auf NEUE
        // DEX_Events-Items (GetOnNewItems). Ein reines MERGE-Update würde
        // den Flow nie anstoßen → kein Outlook-Termin.
        const wasOutlookDisabled = !!initialMeta?.disableOutlook;
        const nowOutlookEnabled = !draft.disableOutlook;
        const hadOutlookEventId = !!(initialMeta?.outlookEventId);
        // v28.69: „Fehlende Termine jetzt anlegen" erzwingt denselben Pfad —
        // ein reines Update triggert den GetOnNewItems-Flow nie.
        const forcedRecreate = forceOutlookRecreateRef.current.has(draft.dbId) && nowOutlookEnabled;
        const needsOutlookRecreate = forcedRecreate || (wasOutlookDisabled && nowOutlookEnabled && !hadOutlookEventId);
        if (needsOutlookRecreate) {
          // v11.69: Seit dem Subsite-Reuse-Pfad muss hier KEINE destruktive
          // Lösch-Aktion mehr passieren. Wir entfernen nur die DEX_Events-
          // Zeile und legen sie mit `existingSubsiteUrl` neu an — alle
          // Anmeldungen, TeilnehmerIDs und die Subsite bleiben unangetastet.
          // Daher auch kein window.confirm mehr nötig.
          const subsiteUrlForReuse = initialMeta?.subsiteUrl || '';
          const regListNameForReuse = initialMeta?.registrationListName || 'Teilnehmer';
          if (subsiteUrlForReuse && regListNameForReuse) {
            // v30.67: Derselbe Riegel wie im v29.21-Pfad oben — dieser
            // ältere Zweig (DisableOutlook-Toggle, „Fehlende Termine jetzt
            // anlegen") war damals übersehen worden. deleteEventItemOnly
            // wirft nie, es meldet Fehlschlag als false; das try/catch war
            // toter Code. Unter Drosselung (429, ab ca. dem 20. Schreib-
            // vorgang) lief der Recreate trotzdem weiter: eine ZWEITE Zeile
            // auf dieselbe Subsite, die alte blieb stehen, fiel nicht in
            // keptDbIds und wurde am Ende KASKADIEREND gelöscht — samt der
            // geteilten Subsite mit allen Anmeldungen.
            // v30.67 (Review): s. recreateWithReuse — erst anlegen, dann
            // löschen. Kein continue bei false: unten in den normalen
            // Update-Pfad fallen, dort landet die dbId in keptDbIds.
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const reusePayload: any = {
              ...childPayload,
              disableOutlook: false,
              outlookEventId: '',
              existingSubsiteUrl: subsiteUrlForReuse,
              existingRegistrationListName: regListNameForReuse,
              onProgress: subOnProgress,
            };
            if (await recreateWithReuse(draft, reusePayload, 'v11.69/Legacy')) continue;
          } else {
            // Edge-Case: kein subsiteUrl auf dem alten Item bekannt (sehr
            // alte Events). In dem Fall fallen wir auf den destruktiven
            // Legacy-Pfad zurück — mit Confirm.
            const msg = isDe
              ? `Beim Sub-Event „${draft.title}" wurde „Outlook-Termin erstellen" nachträglich aktiviert, aber es konnte keine bestehende Teilnehmer-Subsite ermittelt werden.\n\n`
                + `Wenn du jetzt fortfährst, wird das Sub-Event komplett neu aufgesetzt — vorhandene Anmeldungen, Teilnehmer-IDs und die Teilnehmer-Subsite gehen verloren (landen 93 Tage im Papierkorb).\n\n`
                + `Trotzdem fortfahren?`
              : `On sub-event "${draft.title}" you turned on "Create Outlook event" after the fact, but no existing participant subsite could be determined.\n\n`
                + `If you continue, the sub-event will be re-created from scratch — existing registrations, participant IDs and the participant subsite will be lost (recycled for 93 days).\n\n`
                + `Continue anyway?`;
            const confirmed = await confirmDialog(msg, { danger: true, title: isDe ? 'Sub-Event neu aufsetzen' : 'Recreate sub-event', confirmLabel: isDe ? 'Trotzdem fortfahren' : 'Continue anyway' });
            if (confirmed) {
              // v30.67: Gelungenes Löschen merken — sonst versucht die
              // Aufräum-Schleife unten dieselbe Id ein zweites Mal (404).
              const legacyGone = await deleteEvent(draft.dbId).catch(() => false);
              if (legacyGone) replacedDbIds.add(draft.dbId);
              const recreatedLegacyId = await createEvent({ ...childPayload, onProgress: subOnProgress });
              // v27.11: Sub-Event-Bild auch hier persistieren.
              await persistSubEventImage(recreatedLegacyId, draft);
              continue;
            } else {
              draft.disableOutlook = true;
              childPayload.disableOutlook = true;
            }
          }
        }
        keptDbIds.add(draft.dbId);
        // v29.57: Unverändertes Sub-Event überspringen — spart bei einer
        // Terminreihe die große Mehrheit der Schreibvorgänge (s. Kommentar
        // an subPersistKey). Nur wenn BEIDES unverändert ist: der Entwurf
        // selbst UND alles am Hauptevent, was ein Sub-Event erbt.
        // v30.10: … UND der wirksame Outlook-Betreff schon in der Zeile
        // steht. Der v30.7-Fallback (Kalender-Tag → Hauptevent-Titel) wird
        // erst BEIM SCHREIBEN berechnet — im Entwurf ändert sich dadurch
        // nichts, der Skip hielt Bestands-Terminreihen deshalb für
        // unverändert, OutlookSubject blieb leer und der UpdateEvent-Flow
        // fiel per coalesce weiter auf den Tages-Titel zurück, egal wie oft
        // aktualisiert wurde. Genau die Falle, vor der der subPersistKey-
        // Kommentar warnt: ein abgeleiteter Wert, den keiner der beiden
        // Schlüssel sieht.
        const storedSubRow = childEventsOf(parentEventId).find(k => k.id === draft.dbId);
        const subjectInSync = subOutlookSubject === ((storedSubRow && storedSubRow.outlookSubject) || '').trim();
        if (subGateUnchanged && subjectInSync && !draft.imageFile && !draft.imageRemoved) {
          const before = initialSubPersistRef.current[draft.dbId];
          if (before !== undefined && before === subPersistKey(draft)) {
            skippedSubCount++;
            stepDone++;
            if (onStep) onStep(stepDone, stepTotal, shortSubEventTitle(draft.title, title) || draft.title);
            continue;
          }
        }
        // Update bestehender Sub-Event: nur geänderte Felder patchen. CustomFields
        // werden als JSON-String serialisiert — v17.22: zentraler
        // serializeCustomFields-Helper, damit der Sub-Event-Pfad dieselben
        // EN-Varianten + Options-Pairing erhält wie Top-Level (vorher droppte
        // dieser Pfad labelEn/helpTextEn/confirmLabelEn/optionsEn still).
        const cfJson = JSON.stringify(serializeCustomFields(draft.customFields || [], bilingualFields));
        // v11.57: Sub-Event-Kommunikations-Felder mit-persistieren — bisher
        // wurden Mail-Sprache, Outlook-Body, Logos pro Sub-Event nur am
        // Top-Level gespeichert. Mit den Tabs in Step 5 kann jeder Sub-Event
        // jetzt seine eigene Konfiguration haben.
        // v11.57: Outlook-Update-Flag pro Sub-Event setzen, wenn dieses
        // Sub-Event als „outlookDirty" markiert wurde (vom Confirm-Modal
        // entschieden).
        const subUpdates: Record<string, unknown> = {
          'Title': childPayload.title,
          'Description': childPayload.description,
          'Location': childPayload.location,
          'StartDate': childPayload.startDate || null,
          // v28.66 BUG-FIX: hier fehlte der v22.17-Schutz — beim Speichern
          // eines BESTEHENDEN Sub-Events ohne End-Zeit wurde EndDate mit null
          // überschrieben (Outlook-Flow-Crash, s.o.). childPayload trägt den
          // Fallback bereits; die zweite Stufe bleibt als Absicherung stehen.
          'EndDate': childPayload.endDate || childPayload.startDate || null,
          'RegistrationDeadline': childPayload.registrationDeadline || null,
          'MaxParticipants': childPayload.maxParticipants,
          // v20.0 BUG-FIX (Audit): diese Felder wurden beim UPDATE bestehender
          // Sub-Events NIE mitgeschrieben (nur beim Create via childPayload) —
          // Änderungen an Sichtbarkeit (v19.27 AudiencePicker), Adresse, Agenda,
          // Transferzeiten, Abmeldefrist, Warteliste und Anrede-Abfrage gingen
          // beim Speichern eines bestehenden Sub-Events still verloren
          // (gleiche Bug-Klasse wie v19.32 bei den Mail-Flags).
          'WaitlistEnabled': childPayload.waitlistEnabled,
          // v29.20 (Audit): Organizer-Aenderungen erreichten bestehende
          // Sub-Events nie — nur der Create-Pfad schrieb sie (childPayload).
          // Sub-Event-Mails und Flows lesen die Zeile des Sub-Events und
          // adressierten nach einem Organizer-Wechsel weiter die alten
          // Personen. Der Doku-Kommentar oben verspricht die Vererbung
          // ausdruecklich („Alle Sub-Events erben Metadaten (Organizer, …)").
          'Organizer': childPayload.organizer,
          'OrganizerEmail': childPayload.organizerEmail,
          'MandatoryRegistration': childPayload.mandatoryRegistration, // v24.64: Pflicht-Sub-Event
          'LastDeregisterDate': childPayload.lastDeregisterDate || null,
          'LocationAddress': childPayload.locationAddress,
          'LocationFilter': childPayload.locationFilter,
          'Audience': childPayload.audience,
          'FilterMode': childPayload.filterMode,
          // v22.10: Ausschluss-Liste pro Sub-Event aktualisieren (semikolon-sep.).
          'ExcludedUsers': (draft.excludedUsers || []).filter(Boolean).join(';'),
          'Agenda': childPayload.agenda,
          'Transfers': childPayload.transfers,
          'AskSalutation': childPayload.askSalutation,
          'DisableEmails': childPayload.disableEmails,
          // v19.32 BUG-FIX: die granularen An-/Abmelde-Mail-Flags + Auto-Abmeldung
          // wurden beim UPDATE bestehender Sub-Events NICHT mitgeschrieben (nur
          // beim Create), daher gingen sie nach dem Speichern verloren.
          'DisableRegistrationEmail': childPayload.disableRegistrationEmail,
          'DisableCancellationEmail': childPayload.disableCancellationEmail,
          'AutoDeregisterOnDecline': childPayload.autoDeregisterOnDecline,
          'InactiveHandling': childPayload.inactiveHandling || 'notify',
          'DisableOutlook': childPayload.disableOutlook,
          'OutlookSubject': subOutlookSubject,
          'OutlookStart': (draft.outlookStart || '') || null,
          'OutlookEnd': (draft.outlookEnd || '') || null,
          'OutlookLocation': (draft.outlookLocation || '') || '',
          // v29.52: ganztägig auch beim UPDATE bestehender Sub-Events — genau
          // die Klasse Fehler, die v19.32/v20.0/v29.20 schon dreimal hatten.
          'AllDay': !!draft.allDay,
          'ShowAsFree': !!draft.showAsFree, // v29.54
          'OutlookIsOnlineMeeting': onlineMeetingMode === 'auto', // v30.26
          'SkipOrganizerInvite': !orgGetsSubInvites, // v29.55
          'EmailLanguage': childPayload.emailLanguage,
          // v29.42: Fußzeilen-Link auch auf dem direkten Sub-Event-Schreibweg
          // normalisieren (der läuft nicht über EventService.updateEvent).
          'OutlookBody': normalizeMadeWithLink(childPayload.outlookBody || ''),
          'EmailTemplateOverrides': childPayload.emailTemplateOverrides,
          'EmailImageBase64': subEmailLogo || '',
          'CustomFields': cfJson,
        };
        // v22.15: Auto-Heilung — steht das Sub-Event auf „Abgeschlossen"
        // (z.B. vom Auto-Cleanup wegen eines alten Testdatums), das End-Datum
        // liegt nach dieser Bearbeitung aber in der Zukunft, zurück auf Aktiv.
        {
          const storedChild = childEventsOf(parentEventId).find(k => k.id === draft.dbId);
          const subEndIso = childPayload.endDate || childPayload.startDate || '';
          if (storedChild && storedChild.status === 'Completed' && subEndIso && new Date(subEndIso).getTime() > Date.now()) {
            subUpdates['EventStatus'] = 'Active';
          }
        }
        // OutlookDirty + Update wird vom Aufrufer (handleSubmit) anhand des
        // jeweiligen Sub-Event-Snapshots gesteuert — siehe pendingSubUpdates.
        // v29.21 (Audit): Ergebnis PRÜFEN — updateEvent wirft nie, es meldet
        // Fehlschlag als false. Vorher lief der Balken weiter und der Save
        // endete mit „Änderungen gespeichert!", auch wenn einzelne Sub-Events
        // (429, zu großer Payload) nie ankamen.
        // v29.77: skipReload — sonst laedt JEDES Sub-Event-Update die komplette
        // Event-Liste (28 MB) neu. Der eine Reload kommt vom Hauptevent-Update.
        const subOk = await updateEvent(draft.dbId, subUpdates, { skipReload: true });
        if (!subOk) failedSubTitles.push(shortSubEventTitle(draft.title, title) || draft.title);
        // v29.74: Bei Drosselung ANHALTEN statt weiterhaemmern. Zwei
        // Fehlschlaege in Folge waehrend aktiver Drossel-Schranke heisst:
        // SharePoint will gerade keine weiteren Schreibzugriffe von diesem
        // Nutzer. Die restlichen Termine trotzdem zu versuchen (jeder mit
        // eigenen Retries) hat einen Organizer bis zur NUTZER-SPERRE
        // (Throttle.htm) eskaliert. Lieber ehrlich abbrechen — gespeichert
        // ist gespeichert, der Rest kommt beim naechsten Save.
        if (!subOk) {
          consecutiveSubFailures++;
          if (consecutiveSubFailures >= 2 && isThrottled()) {
            abortedForThrottle = true;
            break;
          }
        } else {
          consecutiveSubFailures = 0;
        }
        // v29.74: Atempause zwischen den Schreibzugriffen — 19 MERGEs im
        // Renn-Tempo sind genau das Muster, das die Drosselung ausloest.
        await new Promise<void>(r => setTimeout(r, 250));
        // v27.11: eigenes Sub-Event-Bild persistieren (Upload/Entfernen).
        await persistSubEventImage(draft.dbId, draft);
      } else {
        // v30.67 (Review): createEvent WIRFT seit v30.67 (EventNumber nicht
        // lesbar, 2-MB-Grenze) — ungefangen brach das die ganze Schleife ab:
        // die restlichen Termine und die Aufräum-Schleife liefen nicht, und
        // wizardSubmit meldete nach einem console.warn trotzdem Erfolg.
        let newSubId: number | null = null;
        try { newSubId = await createEvent({ ...childPayload, onProgress: subOnProgress }); }
        catch (err) { console.warn('[DEX] Sub-Event anlegen fehlgeschlagen:', draft.title, err); newSubId = null; }
        if (!newSubId) failedSubTitles.push(shortSubEventTitle(draft.title, title) || draft.title);
        // v27.11: Bild fürs frisch angelegte Sub-Event hochladen (braucht die
        // neue DEX_Events-Item-Id aus createEvent).
        await persistSubEventImage(newSubId, draft);
      }
      stepDone++;
      if (onStep) onStep(stepDone, stepTotal, shortSubEventTitle(draft.title, title) || draft.title);
    }
    // Entfernte Sub-Events aufräumen: deleteEvent löscht kaskadierend auch
    // die Subsite (Teilnehmerliste) und queued einen Outlook-DeleteEvent.
    //
    // v29.48 BUG-FIX: Das Ergebnis wurde hier verworfen („Delete-Fehler darf
    // Save nicht blockieren") — deleteEvent WIRFT aber gar nicht, es liefert
    // false. Ein abgelehntes Löschen (typisch: HTTP 429, SharePoint drosselt
    // nach 20 Sub-Event-Schreibvorgängen) war damit unsichtbar: der Wizard
    // meldete „Änderungen gespeichert!", und der abgewählte Tag stand nach dem
    // Neuladen wieder in der Liste. Genau das war der Kalender-Fall aus der
    // Rückmeldung — 28.09. und 30.09. blieben stehen, obwohl sie abgewählt
    // waren, während der dazwischenliegende 29.09. verschwand.
    const failedDeleteTitles: string[] = [];
    for (const oldId of initialSubEventDbIds) {
      // v30.67: ersetzte Ids überspringen (s. replacedDbIds oben).
      if (!keptDbIds.has(oldId) && !replacedDbIds.has(oldId)) {
        const gone = await deleteEvent(oldId).catch(() => false);
        if (!gone) {
          const stored = childEventsOf(parentEventId).find(k => k.id === oldId);
          failedDeleteTitles.push(stored ? (shortSubEventTitle(stored.title, title) || stored.title) : `#${oldId}`);
        }
      }
    }
    // v29.21 (Audit): gescheiterte Sub-Events benennen statt still erfolgreich
    // zu wirken — der Organizer entscheidet dann selbst, ob er erneut speichert.
    if (abortedForThrottle) {
      showAlert(isDe
        ? `SharePoint bremst gerade alle Schreibzugriffe (Drosselung). Der Speichervorgang wurde nach ${stepDone} von ${stepTotal} Terminen angehalten, damit dein Konto nicht gesperrt wird. Was gespeichert ist, bleibt gespeichert — bitte warte ein paar Minuten und speichere dann erneut; bereits gesicherte Termine werden dabei übersprungen.`
        : `SharePoint is currently throttling all writes. Saving stopped after ${stepDone} of ${stepTotal} dates to protect your account from being blocked. Everything saved so far is kept — please wait a few minutes and save again; dates already saved will be skipped.`, { variant: 'error' });
    } else if (failedSubTitles.length > 0) {
      showAlert(isDe
        ? `${failedSubTitles.length} Sub-Event${failedSubTitles.length === 1 ? '' : 's'} konnte${failedSubTitles.length === 1 ? '' : 'n'} nicht gespeichert werden: ${failedSubTitles.join(', ')}. Die übrigen Änderungen sind gespeichert — bitte speichere erneut, um es nochmal zu versuchen.`
        : `${failedSubTitles.length} sub-event${failedSubTitles.length === 1 ? '' : 's'} could not be saved: ${failedSubTitles.join(', ')}. All other changes are saved — please save again to retry.`, { variant: 'error' });
    }
    // v30.67 (Review): Nach einem Recreate kennt der Client-State die neuen
    // Zeilen nicht — die alten stehen weiter in der Liste, die neuen fehlen,
    // und ein erneut geöffneter Wizard böte dieselben Termine noch einmal an
    // (dann mit dem falschen Grund „Austausch abgelehnt", weil die alte Id
    // längst 404 ist). Bis v30.66 lud die Aufräum-Schleife je ersetzter Id
    // komplett nach — fälschlich, aber es war das einzige Nachladen. Seit sie
    // ersetzte Ids überspringt: hier EINMAL nachladen.
    if (replacedDbIds.size > 0) {
      try { await refreshEvents(); }
      catch (err) { console.warn('[DEX][v30.67] Reload nach Recreate fehlgeschlagen:', err); }
    }
    // v29.57: Nach dem Save ist der aktuelle Stand der neue Vergleichspunkt —
    // sonst würde ein zweiter Save in derselben Sitzung alles erneut schreiben.
    {
      const map: Record<string, string> = {};
      for (const se of subEventsRef.current) if (se.dbId) map[se.dbId] = subPersistKey(se);
      initialSubPersistRef.current = map;
      subTopGateInitialRef.current = subTopGateKey();
      if (skippedSubCount > 0) {
        // eslint-disable-next-line no-console
        dlog('perf', `[DEX] ${skippedSubCount} unveränderte Sub-Events übersprungen (kein Schreibvorgang).`);
      }
    }
    if (failedDeleteTitles.length > 0) {
      showAlert(isDe
        ? `${failedDeleteTitles.length} abgewählte${failedDeleteTitles.length === 1 ? 'r Termin konnte' : ' Termine konnten'} nicht gelöscht werden: ${failedDeleteTitles.join(', ')}. ${failedDeleteTitles.length === 1 ? 'Er steht' : 'Sie stehen'} deshalb weiterhin in der Liste — bitte speichere erneut.`
        : `${failedDeleteTitles.length} deselected date${failedDeleteTitles.length === 1 ? '' : 's'} could not be deleted: ${failedDeleteTitles.join(', ')}. ${failedDeleteTitles.length === 1 ? 'It is' : 'They are'} therefore still in the list — please save again.`, { variant: 'error' });
    }
}
