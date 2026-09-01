/**
 * SP-Roh-Event -> `DeloitteEvent` (Anzeige-Modell der App).
 *
 * v30.66: Aus `EventContext.tsx` herausgezogen (Modularisierung Stufe 3).
 * Die Funktion hing im Provider nur an EINER Stelle am State: sie merkt sich
 * die Subsite-URL je Event. Genau die wird jetzt als `subsiteMap`-Ref
 * hereingereicht — sonst ist der Koerper unveraendert.
 */

import { DeloitteEvent, DexHotel, DexHotelStay, DexHotelRules } from '../types';
import { SPEvent, CustomField } from '../services/EventService';
import { buildDisplayImageUrl, stripSpNoteWrapper } from './eventTextHelpers';

export async function mapSPEventToDeloitteEvent(e: SPEvent, subsiteMap: { current: Record<string, string> }): Promise<DeloitteEvent> {
  // SubsiteUrl merken
  if (e.SubsiteUrl) {
    subsiteMap.current[e.Id.toString()] = e.SubsiteUrl;
  }

  // Teilnehmeranzahl: v26.63 aus der persistierten DEX_Events-Spalte
  // CurrentParticipants als Startwert (statt hart 0) — so ist die Zahl auch
  // ohne Subsite-Scan verfügbar. loadParticipantCountsForEvents überschreibt
  // sie mit der frischen Zahl, sobald ein Event geöffnet/geladen wird.
  const currentParticipants = (typeof e.CurrentParticipants === 'number') ? e.CurrentParticipants : 0;
  const waitlistCount = 0;

  // Custom Fields parsen
  let customFields: CustomField[] = [];
  try {
    if (e.CustomFields) customFields = JSON.parse(e.CustomFields);
  } catch { /* ungültig */ }
  // v29.77: Der v11.18-Debug-Trace („Raw CustomFields for event …") ist
  // entfernt — er druckte bei JEDEM loadEvents die CustomFields fremder
  // Events in die Konsole (las sich wie fremde Daten im falschen Event)
  // und stringifizierte megabyteweise JSON.

  return {
    id: e.Id.toString(),
    eventNumber: e.EventNumber || 0,
    title: e.Title || '',
    // v5.2: EventType-Spalte deprecated. Typ aus CustomFields ableiten
    // (Fallback auf alten SP-Wert wenn noch vorhanden).
    type: (e.EventType as DeloitteEvent['type'])
      || (customFields.some(f => f.id === 'b2run_startblock') ? 'B2Run' : 'Other'),
    // v11.89: 'Under Construction' aus Legacy-Daten transparent auf 'Active'
    // mappen — der Entwurfs-Zustand lebt jetzt auf IsFictive.
    status: (e.EventStatus === 'Under Construction' ? 'Active' : (e.EventStatus as DeloitteEvent['status'])) || 'Active',
    organizers: (stripSpNoteWrapper(e.Organizer) || '').split(';').map((s: string) => s.trim()).filter((s: string) => s),
    organizerEmails: (stripSpNoteWrapper(e.OrganizerEmail) || '').split(';').map((s: string) => s.trim()).filter((s: string) => s),
    // v10.16: Optionaler Ansprechpartner. ContactInfo ist Note-Feld, daher
    // strippen — Name/Email sind Single-Line, kein Wrapper.
    contactName: e.ContactName || '',
    // v28.5: als Rückfragen-Kontakt markierter Organizer (v26.18-Spalte,
    // Wizard-UI + Anzeige kamen erst mit v28.5).
    contactOrganizerEmail: e.ContactOrganizerEmail || '',
    contactEmail: e.ContactEmail || '',
    contactInfo: stripSpNoteWrapper(e.ContactInfo),
    location: e.Location || '',
    locationAddress: (() => {
      try {
        if (!e.LocationAddress) return undefined;
        const o = JSON.parse(e.LocationAddress);
        return { street: o.street || '', houseNo: o.houseNo || '', zip: o.zip || '', city: o.city || '' };
      } catch { return undefined; }
    })(),
    locationAudience: e.LocationFilter ? e.LocationFilter.split(',').map(s => s.trim()) : [],
    audienceFilter: e.Audience ? e.Audience.split(',').map(s => s.trim()) : [],
    // v16.4: vor-aufgelöste Member-E-Mails der Audience-DLs.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    audienceResolvedEmails: ((e as any).AudienceResolvedEmails || '')
      .split(';').map((s: string) => s.trim().toLowerCase()).filter(Boolean),
    filterMode: (e.FilterMode as 'AND' | 'OR') || 'OR',
    startDate: e.StartDate || '',
    endDate: e.EndDate || '',
    // v26.22: SP-Erstell-Zeitstempel (für die Duplikat-Anzeige „erstellt am …").
    created: (e as { Created?: string }).Created || '',
    registrationDeadline: e.RegistrationDeadline || '',
    lastDeregisterDate: e.LastDeregisterDate || '',
    description: e.Description || '',
    maxParticipants: e.MaxParticipants || 0,
    waitlistEnabled: e.WaitlistEnabled !== false, // default true wenn null/undefined
    mandatoryRegistration: e.MandatoryRegistration === true, // v24.64

    autoSendQRCode: e.AutoSendQRCode === true, // v9.15 — explizites opt-in pro Event
    activeFrom: e.ActiveFrom || undefined, // v9.21 — Auto-Activate-Datum
    currentParticipants,
    waitlistCount,
    imageUrl: buildDisplayImageUrl(e.EventImageUrl || '', (e as { Modified?: string }).Modified),
    subsiteUrl: e.SubsiteUrl || '',
    outlookBody: e.OutlookBody || '',
    outlookSubject: e.OutlookSubject || undefined,
    // v29.52: ganztägiger Termin — der Outlook-Flow macht daraus isAllDay.
    allDay: !!e.AllDay,
    showAsFree: !!e.ShowAsFree,
    outlookIsOnlineMeeting: !!e.OutlookIsOnlineMeeting, // v30.26
    skipOrganizerInvite: !!e.SkipOrganizerInvite,
    outlookStart: e.OutlookStart || undefined,
    outlookEnd: e.OutlookEnd || undefined,
    outlookLocation: e.OutlookLocation || undefined,
    outlookEventId: e.OutlookEventId || '',
    // v11.61: CalendarLink (iCalUId) muss in den Event-Type, weil der
    // DEX_CreateOutlookEvent-Flow nur dieses Feld auf Erfolg setzt — die
    // v11.57-Modal-Erkennung hatte auf OutlookEventId geprüft (immer leer)
    // und das Outlook-Update-Confirm-Modal kam deshalb nie.
    calendarLink: e.CalendarLink || '',
    emailLanguage: e.EmailLanguage || 'EN',
    // v18.35: erzwungene Anmeldeseiten-Sprache (nur 'de'/'en' gültig, sonst undefined).
    registrationLanguage: (e.RegistrationLanguage === 'de' || e.RegistrationLanguage === 'en') ? e.RegistrationLanguage : undefined,
    emailTemplateOverrides: e.EmailTemplateOverrides || '',
    disableEmails: !!e.DisableEmails,
    disableRegistrationEmail: !!e.DisableRegistrationEmail,
    disableCancellationEmail: !!e.DisableCancellationEmail,
    autoDeregisterOnDecline: !!e.AutoDeregisterOnDecline,
    inactiveHandling: e.InactiveHandling === 'autoderegister' ? 'autoderegister' : 'notify',
    disableOutlook: !!e.DisableOutlook,
    // v14.5: requireSubEventSelection als Piggyback im EmailTemplateOverrides-
    // JSON (kein neues SP-Feld nötig).
    // v14.8: subEventsOnlyMode + childEventTerm zusätzlich aus dem
    // Piggyback-Blob auslesen. subEventsOnlyMode impliziert
    // requireSubEventSelection (auch wenn der Flag nicht explizit gesetzt
    // ist).
    requireSubEventSelection: ((): boolean => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return !!(ov && (ov._requireSubEventSelection || ov._subEventsOnlyMode));
      } catch { return false; }
    })(),
    subEventsOnlyMode: ((): boolean => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return !!(ov && ov._subEventsOnlyMode);
      } catch { return false; }
    })(),
    // v29.25: Selbst-Abmeldung komplett deaktiviert bzw. nach der
    // Abmeldefrist gesperrt (Piggybacks _noSelfCancel /
    // _noCancelAfterDeadline). Auswertung in utils/cancelPolicy.
    noSelfCancel: ((): boolean => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return !!(ov && ov._noSelfCancel);
      } catch { return false; }
    })(),
    noCancelAfterDeadline: ((): boolean => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return !!(ov && ov._noCancelAfterDeadline);
      } catch { return false; }
    })(),
    // v28.97: Nur EIN Sub-Event waehlbar (Piggyback _subEventSingleChoice).
    subEventSingleChoice: ((): boolean => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return !!(ov && ov._subEventSingleChoice);
      } catch { return false; }
    })(),
    // v28.91: Sub-Events sind Termine → Kalender-Auswahl (Piggyback
    // _subEventCalendar). Ohne Flag bleibt die Anmeldeseite bei der Liste.
    // v29.67: Freischalt-Regel der Kalender-Termine (s. types/index.ts).
    subEventOpenRule: ((): { mode: 'day' | 'week' | 'fixed'; days?: number; date?: string } | undefined => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        const r = ov && ov._subEventOpenRule;
        if (r && (r.mode === 'day' || r.mode === 'week') && typeof r.days === 'number' && r.days > 0) {
          return { mode: r.mode, days: r.days };
        }
        // v29.76: festes Datum — alle Termine oeffnen gemeinsam.
        if (r && r.mode === 'fixed' && typeof r.date === 'string' && r.date) {
          return { mode: 'fixed', date: r.date };
        }
      } catch { /* */ }
      return undefined;
    })(),
    subEventCalendar: ((): boolean => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return !!(ov && ov._subEventCalendar);
      } catch { return false; }
    })(),
    // v28.5: Event-Bild als Banner über den Infos (Piggyback _imageBanner).
    imageBanner: ((): boolean => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return !!(ov && ov._imageBanner);
      } catch { return false; }
    })(),
    // v28.11: URL des UNBESCHNITTENEN Querformat-Originals (Piggyback
    // _imageOrigUrl) — nur gesetzt, wenn ein Querformat-Foto per App-
    // Zuschnitt rund/quadratisch wurde. Die Event-Liste nutzt dann das
    // Original als Kachel-Hintergrund; die Anmeldeseite behält den Kreis.
    imageOrigUrl: ((): string => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return (ov && typeof ov._imageOrigUrl === 'string') ? ov._imageOrigUrl : '';
      } catch { return ''; }
    })(),
    // v29.13: Das Mail-Logo (Piggyback `_eventLogo`, identisch mit der Spalte
    // EmailImageBase64) als Base64. Es ist ein ANDERES Bild als das
    // Event-Bild aus Schritt 1: Mails und Outlook-Termin zeigen dieses,
    // Anmeldeseite und Kachel das Event-Bild. Wer nur eines von beiden
    // pflegt — meist das Mail-Logo, weil man es im Postfach sofort sieht —
    // bekam auf der Anmeldeseite den generischen DEX-Kreis. Deshalb steht
    // es hier zur Verfügung und dient dort als Rückfall.
    mailImageBase64: ((): string => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        const v = (ov && typeof ov._eventLogo === 'string') ? ov._eventLogo : '';
        return v.indexOf('data:') === 0 ? v : '';
      } catch { return ''; }
    })(),
    // v29.38: Vom Organizer hinterlegter Teams-Link (Piggyback `_teamsLink`).
    teamsLink: ((): string => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        const v = (ov && typeof ov._teamsLink === 'string') ? ov._teamsLink.trim() : '';
        return /^https?:\/\//i.test(v) ? v : '';
      } catch { return ''; }
    })(),
    // v28.38: Hotel-Planung (Piggybacks _hotels / _hotelStays / _hotelVisible).
    // Nur Stammdaten und Vorlagen — die Zuordnung pro Person steht in der
    // Teilnehmerliste (Spalten Hotel/HotelFrom/HotelTo).
    hotels: ((): DexHotel[] => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return (ov && Array.isArray(ov._hotels)) ? ov._hotels as DexHotel[] : [];
      } catch { return []; }
    })(),
    hotelStays: ((): DexHotelStay[] => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return (ov && Array.isArray(ov._hotelStays)) ? ov._hotelStays as DexHotelStay[] : [];
      } catch { return []; }
    })(),
    hotelVisibleToAttendees: ((): boolean => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return !!(ov && ov._hotelVisible);
      } catch { return false; }
    })(),
    // v28.58: Verteil-Regeln aus dem Einrichtungs-Assistenten.
    hotelRules: ((): DexHotelRules => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return (ov && ov._hotelRules && typeof ov._hotelRules === 'object') ? ov._hotelRules as DexHotelRules : {};
      } catch { return {}; }
    })(),
    // v28.20: Explizite Klammer-Anmeldefrist (Piggyback _klammerDeadline).
    // Abgelaufen = Gesamt-Event zu (auch bei offenen Sub-Events); leer =
    // wie bisher offen bis zur spätesten Sub-Event-Frist.
    klammerDeadline: ((): string => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return (ov && typeof ov._klammerDeadline === 'string') ? ov._klammerDeadline : '';
      } catch { return ''; }
    })(),
    // v30.6: Rollierende Fristen-Regel als Signal fuer die Anmeldelogik —
    // aktive Reg-Regel = Fristen gelten je Termin, kein harter
    // Klammer-Schluss mehr (isRegistrationFullyClosed).
    subDeadlineRule: (() => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        const r = ov && ov._subDeadlineRule;
        return (r && typeof r === 'object') ? r : undefined;
      } catch { return undefined; }
    })(),
    // v28.2: Sub-Events SOFT-deaktiviert (Piggyback _subEventsDisabled) —
    // die Kind-Events bleiben inkl. Anmeldungen gespeichert, werden aber
    // auf der Anmeldeseite nicht mehr angeboten. Wieder-Einschalten im
    // Wizard zeigt sie unverändert an (kein Löschen mehr über den Toggle).
    subEventsDisabled: ((): boolean => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return !!(ov && ov._subEventsDisabled);
      } catch { return false; }
    })(),
    // v18.9: Organizer-Anzeige ausblenden (Piggyback _hideOrganizer).
    hideOrganizer: ((): boolean => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return !!(ov && ov._hideOrganizer);
      } catch { return false; }
    })(),
    // v24.15: „nur einzelne ausblenden"-Modus (Piggyback _hideOrgIndividual).
    hideOrganizerIndividualOnly: ((): boolean => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return !!(ov && ov._hideOrgIndividual);
      } catch { return false; }
    })(),
    // v24.8 (J): einzelne ausgeblendete Organizer (Piggyback _hiddenOrganizers).
    hiddenOrganizerEmails: ((): string[] => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        const arr = ov && ov._hiddenOrganizers;
        return Array.isArray(arr) ? arr.map((x: unknown) => String(x || '').toLowerCase()).filter(Boolean) : [];
      } catch { return []; }
    })(),
    // v23.6: Assistenz-Sichtbarkeit (Piggyback _assistantsCanSee).
    assistantsCanSee: ((): boolean => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return !!(ov && ov._assistantsCanSee);
      } catch { return false; }
    })(),
    // v23.14: Vorschau vor Aktivierung (Piggyback _previewBeforeActive).
    previewBeforeActive: ((): boolean => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return !!(ov && ov._previewBeforeActive);
      } catch { return false; }
    })(),
    // v23.25: Organizer groß (Foto + Mail direkt sichtbar) statt klein
    // (Chip mit Hover) — Piggyback _organizerDisplayLarge.
    organizerDisplayLarge: ((): boolean => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return !!(ov && ov._organizerDisplayLarge);
      } catch { return false; }
    })(),
    // v23.19: Pro-Ansicht-Darstellung des Event-Bildes (Piggyback _imageDisplay).
    imageDisplay: ((): { card?: { zoom: number; posY: number }; hero?: { zoom: number; posY: number } } | undefined => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        const d = ov && ov._imageDisplay;
        if (d && typeof d === 'object' && (d.card || d.hero)) return d;
        return undefined;
      } catch { return undefined; }
    })(),
    childEventTermSingular: ((): string | undefined => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        const v = ov && ov._childEventTerm && typeof ov._childEventTerm.singular === 'string' ? ov._childEventTerm.singular : '';
        return v || undefined;
      } catch { return undefined; }
    })(),
    childEventTermPlural: ((): string | undefined => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        const v = ov && ov._childEventTerm && typeof ov._childEventTerm.plural === 'string' ? ov._childEventTerm.plural : '';
        return v || undefined;
      } catch { return undefined; }
    })(),
    // v29.60: Geschlecht der Bezeichnung — fuer den unbestimmten Artikel.
    childEventTermGender: ((): 'm' | 'f' | 'n' | undefined => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        const v = ov && ov._childEventTerm ? ov._childEventTerm.gender : '';
        return (v === 'm' || v === 'f' || v === 'n') ? v : undefined;
      } catch { return undefined; }
    })(),
    // v22.78: frei benennbarer Team-Begriff + „keine neuen Teams"-Flag
    // (Piggyback im EmailTemplateOverrides-JSON, analog _childEventTerm).
    teamTermSingular: ((): string | undefined => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        const v = ov && ov._teamTerm && typeof ov._teamTerm.singular === 'string' ? ov._teamTerm.singular : '';
        return v || undefined;
      } catch { return undefined; }
    })(),
    teamTermPlural: ((): string | undefined => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        const v = ov && ov._teamTerm && typeof ov._teamTerm.plural === 'string' ? ov._teamTerm.plural : '';
        return v || undefined;
      } catch { return undefined; }
    })(),
    teamMembersCannotCreate: ((): boolean => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        return !!(ov && ov._teamMembersCannotCreate);
      } catch { return false; }
    })(),
    // v24.58: Anzeige-Bezeichnung des Haupt-Events in der Sub-Event-Auswahl
    // (Piggyback _mainEventLabel = { mode, text }).
    mainEventLabelMode: ((): 'default' | 'custom' | 'none' | undefined => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        const m = ov && ov._mainEventLabel && ov._mainEventLabel.mode;
        return (m === 'custom' || m === 'none') ? m : undefined;
      } catch { return undefined; }
    })(),
    mainEventLabel: ((): string | undefined => {
      try {
        const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
        const v = ov && ov._mainEventLabel && typeof ov._mainEventLabel.text === 'string' ? ov._mainEventLabel.text : '';
        return v || undefined;
      } catch { return undefined; }
    })(),
    // v11.57: bei alten Tenants kann die SP-Spalte fehlen — undefined wird
    // als false interpretiert (kein Hinweis anzeigen).
    outlookDirty: !!e.OutlookDirty,
    notifyOrgRegisterMode: ((): 'never' | 'always' | 'fromDate' => {
      const v = (e.NotifyOrgRegisterMode || '').toLowerCase();
      if (v === 'always') return 'always';
      if (v === 'fromdate') return 'fromDate';
      return 'never';
    })(),
    notifyOrgRegisterFromDate: e.NotifyOrgRegisterFromDate || '',
    notifyOrgCancelMode: ((): 'never' | 'always' | 'afterDeadline' => {
      const v = (e.NotifyOrgCancelMode || '').toLowerCase();
      if (v === 'always') return 'always';
      if (v === 'afterdeadline') return 'afterDeadline';
      return 'never';
    })(),
    excludedUsers: (e.ExcludedUsers || '').split(';').map((s: string) => s.trim().toLowerCase()).filter(Boolean),
    // v11.89: Legacy-Events mit EventStatus='Under Construction' werden
    // auch ohne explizites IsFictive-Flag als Entwurf erkannt — bis die
    // Migration im Hintergrund das SP-Item neu geschrieben hat.
    isFictive: !!e.IsFictive || e.EventStatus === 'Under Construction',
    durchstarterCapacity: typeof e.DurchstarterCapacity === 'number' ? e.DurchstarterCapacity : undefined,
    funstarterCapacity: typeof e.FunstarterCapacity === 'number' ? e.FunstarterCapacity : undefined,
    splitLabelA: e.SplitLabelA || undefined,
    splitLabelB: e.SplitLabelB || undefined,
    splitDescA: e.SplitDescA || undefined,
    splitDescB: e.SplitDescB || undefined,
    splitHelpText: e.SplitHelpText || undefined,
    splitSectionTitle: e.SplitSectionTitle || undefined,
    splitSharedWaitlist: !!e.SplitSharedWaitlist,
    allowAttendeeUpload: !!e.AllowAttendeeUpload,
    attendeeUploadHint: e.AttendeeUploadHint || undefined,
    attendeeUploadLabel: e.AttendeeUploadLabel || undefined,
    // v11.80: Anrede-Toggle + Team-Anmelde-Konfiguration durchreichen.
    // Alte Tenants ohne diese Spalten interpretieren undefined als false /
    // 0, das passt zum Default-Verhalten (Anrede aus, Team-Anmeldung aus).
    askSalutation: !!e.AskSalutation,
    confirmDialogEnabled: !!e.ConfirmDialogEnabled,
    confirmDialogMode: e.ConfirmDialogMode || '',
    confirmDialogText: e.ConfirmDialogText || '',
    // v18.33: Self-Check-in. Alte Tenants ohne diese Spalten lesen undefined
    // als false / leer — Self-Check-in bleibt dann schlicht aus.
    selfCheckInEnabled: !!e.SelfCheckInEnabled,
    selfCheckInToken: e.SelfCheckInToken || undefined,
    selfCheckInFrom: e.SelfCheckInFrom || undefined,
    selfCheckInTo: e.SelfCheckInTo || undefined,
    teamRegistrationEnabled: !!e.TeamRegistrationEnabled,
    teamSize: typeof e.TeamSize === 'number' ? e.TeamSize : 0,
    askTeamName: !!e.AskTeamName,
    // v11.81: Erweiterte Team-Anmelde-Konfiguration (Beitritts-Modus).
    // Alte Tenants ohne diese Spalten interpretieren undefined als false
    // — das deckt sich mit dem konservativen Default „Nur komplette Teams,
    // keine offenen Slots, keine Approval-Queue".
    teamPartialAllowed: !!e.TeamPartialAllowed,
    teamOpenSlotsVisible: !!e.TeamOpenSlotsVisible,
    teamJoinRequiresApproval: !!e.TeamJoinRequiresApproval,
    // v17.20: Bilingual-Toggle für Custom-Fields (DE + EN).
    bilingualFields: !!e.BilingualFields,
    // v6.15: Extra-B2Run-Config aus EmailTemplateOverrides._b2run (piggyback in
    // der bestehenden JSON-Struktur, keine neue SP-Spalte nötig).
    // v6.19: QR-Code-Scanner-Liste aus EmailTemplateOverrides._qrScanners (piggyback).
    // v9.18: Co-Organizer-Liste aus EmailTemplateOverrides._coOrganizers (piggyback, gleicher Pattern).
    // v9.21: Test-Team-Liste aus EmailTemplateOverrides._testTeam (per-Event statt global).
    ...(() => {
      try {
        const parsed = JSON.parse(e.EmailTemplateOverrides || '{}');
        if (!parsed || typeof parsed !== 'object') return { qrScannerNames: [], qrScannerEmails: [], coOrganizerNames: [], coOrganizerEmails: [], testTeamNames: [], testTeamEmails: [] };
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const b = (parsed as any)._b2run;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const qr = (parsed as any)._qrScanners;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const co = (parsed as any)._coOrganizers;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const tt = (parsed as any)._testTeam;
        // v11.25: pure Display-Reihenfolge-Umkehr für Split-Capacity-Karten.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const splitDispRev = !!(parsed as any)._splitDisplayOrderReversed;
        const b2Part = b && typeof b === 'object' ? {
          durchstarterStartblock: typeof b.durchstarterStartblock === 'string' ? b.durchstarterStartblock : undefined,
          funstarterStartblock: typeof b.funstarterStartblock === 'string' ? b.funstarterStartblock : undefined,
          durchstarterRequiresProof: !!b.durchstarterRequiresProof,
        } : {};
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const qrNames: string[] = Array.isArray(qr) ? qr.map((x: any) => String(x?.name || '')) : [];
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const qrEmails: string[] = Array.isArray(qr) ? qr.map((x: any) => String(x?.email || '')) : [];
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const coNames: string[] = Array.isArray(co) ? co.map((x: any) => String(x?.name || '')) : [];
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const coEmails: string[] = Array.isArray(co) ? co.map((x: any) => String(x?.email || '')) : [];
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ttNames: string[] = Array.isArray(tt) ? tt.map((x: any) => String(x?.name || '')) : [];
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ttEmails: string[] = Array.isArray(tt) ? tt.map((x: any) => String(x?.email || '')) : [];
        return { ...b2Part, splitDisplayOrderReversed: splitDispRev, qrScannerNames: qrNames, qrScannerEmails: qrEmails, coOrganizerNames: coNames, coOrganizerEmails: coEmails, testTeamNames: ttNames, testTeamEmails: ttEmails };
      } catch { return { qrScannerNames: [], qrScannerEmails: [], coOrganizerNames: [], coOrganizerEmails: [], testTeamNames: [], testTeamEmails: [] }; }
    })(),
    agenda: (() => { try { return e.Agenda ? JSON.parse(e.Agenda) : []; } catch { return []; } })(),
    transferTimes: (() => { try { return e.Transfers ? JSON.parse(e.Transfers) : []; } catch { return []; } })(),
    quiz: (() => { try { return e.FunZone ? JSON.parse(e.FunZone) : []; } catch { return []; } })(),
    quizClusterSize: typeof e.QuizClusterSize === 'number' && e.QuizClusterSize >= 1 ? e.QuizClusterSize : undefined,
    parentEventId: e.ParentEventId || undefined,
    documents: [], // Wird per loadAttachments nachgeladen
    eventSpecificFields: customFields.map(cf => ({
      id: cf.id,
      label: cf.label,
      type: cf.type,
      required: cf.required,
      options: cf.options,
      // v7.20: helpText durchreichen, damit das Registrierungsformular ihn
      // im "i"-Tooltip neben dem Label anzeigen kann.
      helpText: cf.helpText || '',
      // v18.18: Darstellungs-Stil der Beschreibung (tooltip|inline).
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      helpTextStyle: (cf as any).helpTextStyle === 'inline' ? 'inline' : 'tooltip',
      // v7.21: showIf-Bedingung durchreichen — RegistrationPage filtert
      // anhand davon, ob das Feld angezeigt wird.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      showIf: (cf as any).showIf,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      spInternalName: (cf as any).spInternalName || '',
      // v7.11: multi-Flag durchreichen, damit RegistrationPage Mehrfachauswahl rendern kann
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      multi: !!(cf as any).multi,
      // v26.74: Vorauswahl (Single-Select) durchreichen — RegistrationPage
      // belegt das Feld damit vor, der Wizard zeigt sie im Feld-Editor.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      defaultValue: (cf as any).defaultValue || undefined,
      // v26.75: Vorfilter-Kategorien + Beschriftung durchreichen.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      optionCategories: Array.isArray((cf as any).optionCategories) ? (cf as any).optionCategories : undefined,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      prefilterLabel: (cf as any).prefilterLabel || undefined,
      // v24.25: withTime — bei Datums-Feldern (type='date') auch die Uhrzeit
      // mit abfragen (datetime-local statt date).
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      withTime: !!(cf as any).withTime,
      // v29.20 (Audit A3): Die daterange-Grenzen (v28.63) fehlten in diesem
      // Mapping komplett — das buchbare Übernachtungs-Fenster und das
      // Nächte-Limit kamen deshalb bei KEINEM Event auf der Anmeldeseite an,
      // und der Wizard konnte sie beim Edit nicht laden (→ der nächste Save
      // entfernte sie endgültig aus dem CustomFields-JSON).
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      rangeStart: (cf as any).rangeStart || undefined,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      rangeEnd: (cf as any).rangeEnd || undefined,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      maxNights: typeof (cf as any).maxNights === 'number' ? (cf as any).maxNights : undefined,
      // externalLinks ebenfalls durchreichen, damit AGB-Links für B2Run-Datenschutz
      // korrekt unter dem Feld angezeigt werden (war bisher nur über den Fallback in
      // RegistrationPage abgesichert).
      externalLinks: cf.externalLinks,
      // v18.41: CC-bei-Mail-Flag durchreichen — collectCcEmailsFromFields liest es.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      ccOnEmails: !!(cf as any).ccOnEmails,
      // v26.60: Roommate-Benachrichtigung — nur explizites false schaltet ab.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      notifyRoommate: (cf as any).notifyRoommate !== false,
      // v11.16: onlyForGroup aus dem persistierten Feld durchreichen.
      // Wurde im Wizard sauber gespeichert (CustomFields-JSON enthält
      // den Schlüssel), aber der Loader hat ihn nie zurückgelesen —
      // Folge: die Gruppen-spezifische Sichtbarkeit (Funstarter only /
      // Durchstarter only) hat in der Registrierungs-UI nie gegriffen,
      // weil die Filter-Chain auf undefined gefallen ist.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      onlyForGroup: (cf as any).onlyForGroup,
      // v17.19: confirmLabel (Text neben Checkbox) im Mapping nachgezogen
      // — vorher hier vergessen, Folge: Wizard speicherte den Text korrekt,
      // RegistrationPage fiel aber immer auf den Default „Ja, bestätigen"
      // zurück, weil das Field-Mapping confirmLabel droppte.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      confirmLabel: (cf as any).confirmLabel,
      // v17.20: Englische Varianten durchreichen — nur dann wirksam, wenn
      // auf Event-Ebene `bilingualFields=true` ist; die RegistrationPage
      // entscheidet zur Laufzeit, ob sie die EN-Spalte zieht.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      labelEn: (cf as any).labelEn,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      helpTextEn: (cf as any).helpTextEn,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      confirmLabelEn: (cf as any).confirmLabelEn,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      optionsEn: (cf as any).optionsEn,
    })),
  };
}
