/**
 * v30.66 — Modularisierung Stufe 2: Thema „Schema der DEX_Events-Liste":
 * Liste und Spalten anlegen, fehlende Spalten nachziehen, Text-Spalten auf
 * Note hochziehen (Überlauf), Berechtigungen und Standard-Ansicht.
 *
 * Die Migrationen sind Bestandspflege: Eine Text-Spalte hält 255 Zeichen, und
 * der 256. schlägt beim Speichern als HTTP 400 auf — deshalb wandern
 * Audience, Organizer & Co. auf Note. Der Umbau löscht die Spalte und legt
 * sie neu an; die Werte werden vorher gesichert und danach zurückgeschrieben.
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';

// ==================== DEX_Events Liste ====================

/**
 * Events-Liste erstellen falls nicht vorhanden
 */
export async function ensureEventsList(svc: EventService): Promise<void> {
  const listName = 'DEX_Events';
  const exists = await svc.listExists(listName);
  if (exists) {
    await ensureMissingFields(svc, listName);

    // Default-View komplett neu aufbauen: ID, Title, EventImageUrl, dann Rest
    try {
      await svc._post(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields/removeallviewfields`,
        {}
      );
    } catch { /* ignore */ }
    await svc.configureDefaultView(listName, [
      'ID', 'LinkTitle', 'EventImageUrl',
      'EventNumber', 'EventStatus', 'Location', 'LocationFilter',
      'StartDate', 'EndDate', 'RegistrationDeadline', 'MaxParticipants',
      'WaitlistEnabled', 'Organizer', 'DisableEmails', 'DisableOutlook',
      'CalendarLink', 'RegistrationListName', 'RegistrationListUrl', 'SubsiteUrl',
    ], undefined, { rebuild: true });
    await setColumnFormatting(svc, listName, 'EventImageUrl', {
      '$schema': 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
      'elmType': 'img',
      'attributes': { 'src': '@currentField' },
      'style': { 'max-height': '60px', 'max-width': '120px', 'border-radius': '6px', 'box-shadow': '0 1px 3px rgba(0,0,0,0.15)' },
    });
    try {
      const listInfo = await svc._sp.get(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=HasUniqueRoleAssignments`,
        SPHttpClient.configurations.v1
      );
      if (listInfo.ok) {
        const data = await listInfo.json();
        if (!data.HasUniqueRoleAssignments) {
          await setEventsListPermissions(svc, listName);
        }
      }
    } catch { /* ignore */ }
    return;
  }

  // Liste erstellen
  await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'Events der DEX Event Experience Platform',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });

  // Spalten hinzufügen
  const fields = getEventsFieldDefinitions();
  for (const f of fields) {
    const payload: Record<string, unknown> = {
      '__metadata': { 'type': f.metaType || 'SP.Field' },
      'Title': f.title,
      'FieldTypeKind': f.type,
      'Required': false,
    };
    if (f.choices) {
      payload['Choices'] = { 'results': f.choices };
    }
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
  }

  await svc.configureDefaultView(listName, [
    'EventNumber', 'EventStatus', 'Location', 'LocationFilter',
    'StartDate', 'EndDate', 'RegistrationDeadline', 'MaxParticipants',
    'WaitlistEnabled', 'Organizer', 'EventImageUrl', 'CalendarLink', 'RegistrationListName', 'RegistrationListUrl', 'SubsiteUrl',
  ], undefined, { rebuild: true });
  await setColumnFormatting(svc, listName, 'EventImageUrl', {
    '$schema': 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
    'elmType': 'img',
    'attributes': { 'src': '@currentField' },
    'style': { 'max-height': '50px', 'max-width': '120px', 'border-radius': '4px' },
  });

  await setEventsListPermissions(svc, listName);
}

/**
 * Feld-Definitionen für DEX_Events Liste
 */
export function getEventsFieldDefinitions(): Array<{ title: string; type: number; choices?: string[]; metaType?: string; richText?: boolean; numberOfLines?: number }> {
  return [
    { title: 'EventStatus', type: 6, choices: ['Under Construction', 'Active', 'Completed', 'Cancelled'], metaType: 'SP.FieldChoice' },
    // EventType-Spalte ab v5.2 deprecated (Feld wird nicht mehr angelegt/aktualisiert).
    // Typ wird beim Laden aus CustomFields abgeleitet. Bestehende Spalte in DEX_Events
    // kann manuell entfernt werden.
    { title: 'Description', type: 3 },
    { title: 'Location', type: 2 },
    { title: 'LocationAddress', type: 2 }, // JSON-String: { street, houseNo, zip, city }
    { title: 'LocationFilter', type: 2 },
    // Audience ist Multi-Line-Text (Note) damit es auch bei 100+ E-Mail-Adressen
    // nicht abgeschnitten wird (Single-Line-Text ist auf 255 Zeichen limitiert).
    // Für bestehende Events siehe upgradeAudienceFieldToNote() — migriert die
    // alte Text-Spalte zu Note ohne Datenverlust.
    { title: 'Audience', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 },
    // v16.4: Vor-aufgelöste E-Mails der Audience-DLs (Multi-Line, ';'-
    // separiert). Wird beim Event-Save vom EventCreationPage-Flow
    // gesetzt; matchesAudience im EventListPage checkt zusätzlich
    // gegen diese Liste. Damit funktioniert die Sichtbarkeit auch für
    // verschachtelte DLs, die /me/memberOf nicht zurückliefert.
    { title: 'AudienceResolvedEmails', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 8 },
    { title: 'FilterMode', type: 6, choices: ['AND', 'OR'], metaType: 'SP.FieldChoice' },
    { title: 'StartDate', type: 4 },
    { title: 'EndDate', type: 4 },
    { title: 'RegistrationDeadline', type: 4 },
    { title: 'LastDeregisterDate', type: 4 },
    { title: 'MaxParticipants', type: 9 },
    // v26.63: Denormalisierte aktuelle Teilnehmerzahl am Event-Item. Wird von
    // Organizern/Admins gepflegt (nur die haben Schreibrechte auf DEX_Events)
    // — beim Laden der echten Zahl best-effort persistiert. So ist die
    // Teilnehmerzahl pro Event ohne Subsite-Scan aus DEX_Events lesbar.
    { title: 'CurrentParticipants', type: 9 },
    { title: 'WaitlistEnabled', type: 8 },
    { title: 'MandatoryRegistration', type: 8 },
    // v26.55: EventImageUrl ist Note (mehrzeilig) statt Single-Line-Text —
    // SharePoint-Asset-URLs (SiteAssets-Pfad + langer Original-Dateiname)
    // überschreiten real das 255-Zeichen-Limit (MD Academy: 261 Zeichen) und
    // ließen den kompletten Event-Save mit „Invalid text value" abbrechen.
    // Bestands-Listen migriert upgradeOverflowTextFieldsToNote() beim Boot
    // bzw. die Selbstheilung in updateEvent() beim ersten fehlschlagenden Save.
    { title: 'EventImageUrl', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 2 },
    { title: 'EmailImageBase64', type: 3 }, // Base64 Event-Bild für E-Mails/Outlook (Flow ersetzt {{ORB_URL}})
    // Organizer + OrganizerEmail sind Multi-Line-Text (Note) damit sie auch bei
    // 10+ Co-Organizern nicht abgeschnitten werden (Single-Line-Text ist auf 255
    // Zeichen limitiert — bei ~17 Personen mit Format `vorname.nachname@deloitte.de;`
    // wird das überschritten und SP antwortet mit „Invalid text value" beim Update).
    // Für bestehende Events siehe upgradeOrganizerFieldsToNote() — migriert die
    // alten Text-Spalten ohne Datenverlust.
    { title: 'Organizer', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 },
    { title: 'OrganizerEmail', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 },
    // v10.16: Optionaler Ansprechpartner pro Event (Anzeige-Feld, kein App-Login).
    { title: 'ContactName', type: 2 },
    { title: 'ContactEmail', type: 2 },
    { title: 'ContactInfo', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 },
    { title: 'EventNumber', type: 9 },
    { title: 'OutlookEventId', type: 2 },
    { title: 'CalendarLink', type: 2 },
    { title: 'OutlookBody', type: 3 }, // Multiline - Text für Outlook-Termin
    { title: 'OutlookSubject', type: 2 }, // v18.42: Single line - Betreff des Outlook-Termins (leer = Titel)
    { title: 'OutlookStart', type: 4 }, // v18.44: DateTime - abweichende Start-Zeit (leer = Event-Start)
    { title: 'OutlookEnd', type: 4 },   // v18.44: DateTime - abweichende End-Zeit (leer = Event-Ende)
    { title: 'OutlookLocation', type: 2 }, // v18.34: Single line - lesbarer Ort für den Outlook-Termin
    // v29.52: Boolean - ganztägiger Termin. Echte Spalte (kein Piggyback in
    // EmailTemplateOverrides), weil der Outlook-Flow den Wert direkt über
    // triggerBody()?['AllDay'] lesen muss — JSON parsen kann er dort nicht.
    { title: 'AllDay', type: 8, metaType: 'SP.Field' },
    // v29.54: Boolean - Termin als „Frei" statt „Beschäftigt" anzeigen.
    // NEGATIV benannt, damit bestehende Einträge (leer/false) weiter
    // beschäftigt bleiben — siehe Kommentar an DeloitteEvent.showAsFree.
    { title: 'ShowAsFree', type: 8, metaType: 'SP.Field' },
    // v30.26: Boolean — Outlook-Termin als echte Teams-Besprechung anlegen
    // (Flow „Create event (V4)": Is online meeting + Provider Teams).
    // Leer/false = wie bisher KEIN automatisches Meeting; der Organizer
    // entscheidet das pro Event im Wizard (Ort → Online-Meeting).
    { title: 'OutlookIsOnlineMeeting', type: 8, metaType: 'SP.Field' },
    // v29.55: Boolean - Organizer nicht in requiredAttendees des Outlook-
    // Termins. Ebenfalls negativ, damit Bestandsevents unverändert bleiben.
    { title: 'SkipOrganizerInvite', type: 8, metaType: 'SP.Field' },
    { title: 'EmailLanguage', type: 2 }, // DE oder EN
    { title: 'RegistrationLanguage', type: 2 }, // v18.35: erzwungene Anmeldeseiten-Sprache ('de'|'en'|'')
    { title: 'EmailTemplateOverrides', type: 3 }, // JSON mit Event-spezifischen Template-Anpassungen
    { title: 'DisableEmails', type: 8, metaType: 'SP.Field' }, // Boolean - keine E-Mails versenden
    { title: 'DisableRegistrationEmail', type: 8, metaType: 'SP.Field' }, // v19.21 Boolean - keine Anmelde-Bestätigung
    { title: 'DisableCancellationEmail', type: 8, metaType: 'SP.Field' }, // v19.21 Boolean - keine Abmelde-Bestätigung
    { title: 'AutoDeregisterOnDecline', type: 8, metaType: 'SP.Field' }, // v19.23 Boolean - Outlook-Absage = Auto-Abmeldung
    { title: 'InactiveHandling', type: 2, metaType: 'SP.Field' }, // v26.40 Text - 'notify' | 'autoderegister' bei Ex-Deloitte-Konten
    { title: 'DisableOutlook', type: 8, metaType: 'SP.Field' }, // Boolean - keine Outlook-Kalendereinträge
    { title: 'OutlookDirty', type: 8, metaType: 'SP.Field' }, // v11.57 Boolean - Outlook-Update ausstehend nach Bearbeitung
    { title: 'AutoSendQRCode', type: 8, metaType: 'SP.Field' }, // v9.15 Boolean - QR-Code automatisch nach Anmeldung versenden
    { title: 'ActiveFrom', type: 4, metaType: 'SP.Field' }, // v9.21 DateTime - Auto-Aktivierungs-Datum
    { title: 'NotifyOrgRegisterMode', type: 6, choices: ['Never', 'Always', 'FromDate'], metaType: 'SP.FieldChoice' }, // v8.5
    { title: 'NotifyOrgRegisterFromDate', type: 4 }, // v8.5: ISO-Date, nur für Mode='FromDate'
    { title: 'NotifyOrgCancelMode', type: 6, choices: ['Never', 'Always', 'AfterDeadline'], metaType: 'SP.FieldChoice' }, // v8.5
    { title: 'ExcludedUsers', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 }, // v8.6: explizit ausgeschlossene User
    { title: 'IsFictive', type: 8, metaType: 'SP.Field' }, // Boolean - Test-Event (nur Admin + eigene Organizer sichtbar)
    { title: 'DurchstarterCapacity', type: 9 }, // Split-Capacity Gruppe A (historisch B2Run-Durchstarter)
    { title: 'FunstarterCapacity', type: 9 }, // Split-Capacity Gruppe B (historisch B2Run-Funstarter)
    { title: 'SplitLabelA', type: 2 }, // v10.20: frei wählbare Bezeichnung Gruppe A (Single line text)
    { title: 'SplitLabelB', type: 2 }, // v10.20: frei wählbare Bezeichnung Gruppe B (Single line text)
    { title: 'SplitDescA', type: 3 }, // v26.72: Beschreibung Gruppe A (Note/mehrzeilig)
    { title: 'SplitDescB', type: 3 }, // v26.72: Beschreibung Gruppe B (Note/mehrzeilig)
    { title: 'SplitHelpText', type: 3 }, // v26.83: Hinweistext über der Gruppen-Auswahl (Note/mehrzeilig)
    { title: 'SplitSectionTitle', type: 2 }, // v26.83: frei wählbare Überschrift der Gruppen-Auswahl (Single line)
    { title: 'SplitSharedWaitlist', type: 8, metaType: 'SP.Field' }, // v10.20: Boolean - true = gemeinsame Warteliste
    { title: 'AllowAttendeeUpload', type: 8, metaType: 'SP.Field' }, // v11.0: Boolean - Teilnehmer-PDF-Upload erlauben
    { title: 'AttendeeUploadHint', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 3 }, // v11.0: Hinweistext
    { title: 'AttendeeUploadLabel', type: 2 }, // v11.0: Single-line Label für den Upload-Block in MyEvents
    { title: 'AskSalutation', type: 8, metaType: 'SP.Field' }, // v11.80: Boolean - Anrede im Registrierungsformular abfragen
    { title: 'ContactOrganizerEmail', type: 2 }, // v26.18: E-Mail des als Hauptkontakt markierten Organizers (grün hervorgehoben auf der Anmeldeseite)
    { title: 'ConfirmDialogEnabled', type: 8, metaType: 'SP.Field' }, // v18.75: Boolean - Sicherheitshinweis vor dem Absenden anzeigen
    { title: 'ConfirmDialogMode', type: 2 }, // v18.75: Single line text - 'summary' (Auswahl-Übersicht) | 'freetext'
    { title: 'ConfirmDialogText', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 }, // v18.75: Note - Freitext-Hinweis
    { title: 'SelfCheckInEnabled', type: 8, metaType: 'SP.Field' }, // v18.33: Boolean - Self-Check-in per QR-Code erlauben
    { title: 'SelfCheckInToken', type: 2 }, // v18.33: Single line text - geheimer Token (statischer Link + HMAC-Schlüssel)
    { title: 'SelfCheckInFrom', type: 4 }, // v18.33: DateTime - optionaler Start des Check-in-Fensters
    { title: 'SelfCheckInTo', type: 4 }, // v18.33: DateTime - optionales Ende des Check-in-Fensters
    { title: 'TeamRegistrationEnabled', type: 8, metaType: 'SP.Field' }, // v11.80: Boolean - Team-Anmeldung erlauben
    { title: 'TeamSize', type: 9 }, // v11.80: Number - Maximale Teamgröße (0 = nicht gesetzt)
    { title: 'AskTeamName', type: 8, metaType: 'SP.Field' }, // v11.80: Boolean - Team-Name abfragen
    { title: 'TeamPartialAllowed', type: 8, metaType: 'SP.Field' }, // v11.81: Boolean - Auch Teil-Teams erlauben
    { title: 'TeamOpenSlotsVisible', type: 8, metaType: 'SP.Field' }, // v11.81: Boolean - offene Slots öffentlich sichtbar
    { title: 'TeamJoinRequiresApproval', type: 8, metaType: 'SP.Field' }, // v11.81: Boolean - Lead muss Beitritt bestätigen
    { title: 'BilingualFields', type: 8, metaType: 'SP.Field' }, // v17.20: Boolean - Custom-Fields zweisprachig (DE + EN) anbieten
    { title: 'CustomFields', type: 3 },
    { title: 'Agenda', type: 3 }, // JSON-Array mit Agenda-Einträgen
    { title: 'Transfers', type: 3 }, // JSON-Array mit Transferzeiten
    { title: 'Documents', type: 3 }, // JSON-Array mit Dokumenten
    { title: 'FunZone', type: 3 }, // JSON-Array mit Quiz-Fragen
    { title: 'QuizClusterSize', type: 9 }, // Number - 1..4 Fragen pro Quiz-Ansicht
    { title: 'ParentEventId', type: 2 }, // Seit v6.4: ID des Parent-Events (wenn dies ein Sub-Event ist)
    { title: 'RegistrationListName', type: 2 },
    { title: 'RegistrationListUrl', type: 2 },
    { title: 'SubsiteUrl', type: 2 },
  ];
}

/**
 * Fehlende Spalten auf einer bestehenden DEX_Events-Liste nachträglich hinzufügen.
 */
async function ensureMissingFields(svc: EventService, listName: string): Promise<void> {
  const requiredFields = getEventsFieldDefinitions();

  try {
    const response = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName&$filter=Hidden eq false&$top=200`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) return;

    const data = await response.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const existingFields = new Set((data.value || []).map((f: any) => f.InternalName));

    for (const f of requiredFields) {
      if (!existingFields.has(f.title)) {
        // Fehlende Spalte nachträglich hinzufügen
        const payload: Record<string, unknown> = {
          '__metadata': { 'type': f.metaType || 'SP.Field' },
          'Title': f.title,
          'FieldTypeKind': f.type,
          'Required': false,
        };
        if (f.choices) {
          payload['Choices'] = { 'results': f.choices };
        }
        if (f.metaType === 'SP.FieldMultiLineText') {
          payload['RichText'] = !!f.richText;
          if (typeof f.numberOfLines === 'number') payload['NumberOfLines'] = f.numberOfLines;
        }
        try {
          await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
        } catch {
          console.warn('[DEX] Konnte Spalte nicht hinzufügen:', f.title);
        }
      }
    }
  } catch (e) {
    console.warn('[DEX] ensureMissingFields Error:', e);
  }
}

/**
 * Migration: alte Audience-Spalte (Type 2, Single-Line-Text, 255 Zeichen Limit)
 * auf Multi-Line-Text (Type 3, Plain-Text, 63.999 Zeichen) umstellen.
 *
 * Nötig weil bei Zielgruppen mit vielen Email-Adressen (~10+) der 255-Zeichen-
 * Cutoff schon griff und Adressen stumm abgeschnitten wurden.
 *
 * Ablauf (idempotent):
 *   1. Check TypeAsString der Audience-Spalte. Wenn schon 'Note' -> skip.
 *   2. Backup aller Event-Werte (id -> audience) im Speicher.
 *   3. Alte Spalte löschen.
 *   4. Neue Spalte als SP.FieldMultiLineText (RichText=false, NumberOfLines=4) anlegen.
 *   5. Werte aus dem Backup zurückschreiben (MERGE pro Event).
 *
 * Läuft beim App-Start (nur für Admins, weil wir Write-Rechte auf DEX_Events brauchen).
 */
export async function upgradeAudienceFieldToNote(svc: EventService): Promise<void> {
  const listName = 'DEX_Events';
  try {
    // 1. TypeAsString abfragen
    const fieldResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Audience')?$select=TypeAsString,FieldTypeKind`,
      SPHttpClient.configurations.v1
    );
    if (!fieldResp.ok) return;
    const fieldData = await fieldResp.json();
    const typeAsString: string = fieldData.TypeAsString || fieldData.d?.TypeAsString || '';
    const fieldTypeKind: number = fieldData.FieldTypeKind ?? fieldData.d?.FieldTypeKind ?? 0;
    if (typeAsString === 'Note' || fieldTypeKind === 3) {
      // Schon migriert
      return;
    }
    if (typeAsString !== 'Text' && fieldTypeKind !== 2) {
      // Unerwarteter Typ - nicht anfassen
      console.warn(`[DEX] upgradeAudienceFieldToNote: Audience hat unerwarteten Typ '${typeAsString}' (kind=${fieldTypeKind}) - skip.`);
      return;
    }

    // 2. Alle Event-Werte laden und backuppen
    const itemsResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Audience&$top=2000`,
      SPHttpClient.configurations.v1
    );
    if (!itemsResp.ok) return;
    const itemsData = await itemsResp.json();
    const items: Array<{ Id: number; Audience: string }> = itemsData.value || itemsData.d?.results || [];
    const backup: Record<number, string> = {};
    for (const it of items) {
      if (it.Audience) backup[it.Id] = it.Audience;
    }
    console.warn(`[DEX] upgradeAudienceFieldToNote: Backup ${Object.keys(backup).length} von ${items.length} Event-Audience-Werten.`);

    // 3. Alte Spalte löschen
    try {
      await svc._post(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Audience')/deleteObject`,
        {}
      );
    } catch (e) {
      console.warn('[DEX] upgradeAudienceFieldToNote: Delete alte Audience-Spalte fehlgeschlagen, Migration abgebrochen:', e);
      return;
    }

    // 4. Neue Spalte als Multi-Line-Text anlegen
    try {
      await svc._post(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
        {
          '__metadata': { 'type': 'SP.FieldMultiLineText' },
          'Title': 'Audience',
          'FieldTypeKind': 3,
          'Required': false,
          'RichText': false,
          'NumberOfLines': 4,
        }
      );
    } catch (e) {
      console.error('[DEX] upgradeAudienceFieldToNote: Konnte neue Audience-Note-Spalte nicht anlegen - Daten könnten verloren gehen:', e, backup);
      return;
    }

    // 5. Werte zurückschreiben per _merge (odata=nometadata, daher kein __metadata im Body nötig)
    let restored = 0;
    let failed = 0;
    for (const idStr of Object.keys(backup)) {
      const id = Number(idStr);
      try {
        const resp = await svc._merge(
          `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`,
          { 'Audience': backup[id] }
        );
        if (resp.ok) restored += 1;
        else failed += 1;
      } catch { failed += 1; }
    }
    console.warn(`[DEX] upgradeAudienceFieldToNote: Migration fertig — ${restored} Werte zurückgeschrieben, ${failed} Fehler.`);
  } catch (e) {
    console.warn('[DEX] upgradeAudienceFieldToNote Error:', e);
  }
}

/**
 * Migration: alte `Organizer` + `OrganizerEmail`-Spalten (Type 2, Single-Line-Text,
 * 255 Zeichen Limit) auf Multi-Line-Text (Type 3, Plain-Text, 63.999 Zeichen) umstellen.
 *
 * Nötig weil bei Events mit 10+ Co-Organizern der 255-Zeichen-Cutoff greift und
 * SharePoint beim Update mit „Invalid text value. A text field contains invalid data."
 * (HTTP 500, Microsoft.SharePoint.SPException) antwortet — der Save bricht komplett ab.
 *
 * Beispiel-Overflow: 17 × `vorname.nachname@deloitte.de;` ≈ 425 Zeichen.
 *
 * Ablauf pro Feld (idempotent, parallel für beide Felder):
 *   1. Check TypeAsString. Wenn schon 'Note' -> skip.
 *   2. Backup aller Event-Werte (id -> wert) im Speicher.
 *   3. Alte Spalte löschen.
 *   4. Neue Spalte als SP.FieldMultiLineText (RichText=false, NumberOfLines=4) anlegen.
 *   5. Werte aus dem Backup zurückschreiben (MERGE pro Event).
 *
 * Läuft beim App-Start (nur für Admins, weil wir Write-Rechte auf DEX_Events brauchen).
 */
export async function upgradeOrganizerFieldsToNote(svc: EventService): Promise<void> {
  await svc._upgradeTextFieldToNote('DEX_Events', 'Organizer');
  await svc._upgradeTextFieldToNote('DEX_Events', 'OrganizerEmail');
}

/**
 * v26.53/v26.55: Migration überlauf-gefährdeter Single-Line-Text-Spalten auf
 * Multi-Line-Text (Note). Auf Bestands-Listen liegen diese Spalten als
 * einzeiliger Text (255-Zeichen-Limit) — längere Werte ließen den kompletten
 * Event-Save mit „Invalid text value. A text field contains invalid data."
 * (HTTP 500) abbrechen. Konkrete Fälle:
 *  - ConfirmDialogText (v26.53): Freitext des Bestätigungs-Dialogs, z. B.
 *    Stornoregeln mit ~450 Zeichen.
 *  - EventImageUrl (v26.55): SiteAssets-Bild-URL mit langem Original-
 *    Dateinamen, real 261 Zeichen (MD Academy 2026).
 * Idempotent: ist eine Spalte schon Note, passiert nichts. Zusätzlich heilt
 * updateEvent() solche Spalten seit v26.54 auch beim fehlschlagenden Save
 * selbst (Boot-Ensure läuft nur einmal pro Version — s. ENSURE_FLAG_KEY).
 */
export async function upgradeOverflowTextFieldsToNote(svc: EventService): Promise<void> {
  await svc._upgradeTextFieldToNote('DEX_Events', 'ConfirmDialogText');
  await svc._upgradeTextFieldToNote('DEX_Events', 'EventImageUrl');
}

/**
 * v26.57: Best-effort-Check, ob eine Person bereits Zugriff auf die Site
 * hat (mindestens Seiten ansehen). Genutzt vor der „SharePoint-Zugriff
 * benötigt"-Admin-Mail bei internationalen Zielgruppen-Personen — wer schon
 * berechtigt ist, taucht in der Mail nicht mehr auf.
 *
 * Ablauf: LoginName über siteusers auflösen (für Gäste weicht er mit
 * #EXT#-Format von der Mail ab), dann getusereffectivepermissions und das
 * ViewPages-Bit prüfen. Gruppen-basierte Rechte (z. B. „Deloitte DE ALL")
 * löst SharePoint dabei serverseitig mit auf.
 *
 * Rückgabe: true = hat Zugriff · false = sicher kein Zugriff · null = nicht
 * prüfbar (User unbekannt, keine Enumerate-Permissions-Rechte des Aufrufers,
 * Netzwerkfehler). Aufrufer behandeln null wie „kein Zugriff" — lieber
 * einmal zu viel benachrichtigen als eine Freigabe verpassen.
 */
export async function userHasSiteAccess(svc: EventService, email: string): Promise<boolean | null> {
  const mail = (email || '').trim();
  if (!mail) return null;
  try {
    let login = `i:0#.f|membership|${mail.toLowerCase()}`;
    try {
      const resp = await svc._sp.get(
        `${svc.siteUrl}/_api/web/siteusers?$filter=Email eq '${encodeURIComponent(mail.replace(/'/g, "''"))}'&$select=LoginName&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const data = await resp.json();
        const item = (data.value || data.d?.results || [])[0];
        if (item && item.LoginName) login = item.LoginName;
      }
    } catch { /* Fallback auf membership-Claim */ }
    const permResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/getusereffectivepermissions(@u)?@u='${encodeURIComponent(login.replace(/'/g, "''"))}'`,
      SPHttpClient.configurations.v1
    );
    if (!permResp.ok) return null;
    const perm = await permResp.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const raw: any = (perm && (perm as any).GetUserEffectivePermissions) || (perm as any).d?.GetUserEffectivePermissions || perm;
    const low = Number(raw?.Low || 0);
    // ViewPages = 0x20000 — reicht, um die App-Seite zu öffnen.
    return (low & 0x20000) !== 0;
  } catch {
    return null;
  }
}

/**
 * Generischer Helper: migriert ein einzelnes Single-Line-Text-Feld einer Liste auf
 * Multi-Line-Text (Note). Idempotent — wenn das Feld schon Note ist, no-op. Wenn das
 * Feld einen anderen Typ hat (Choice/Number/etc.), no-op mit Warnung.
 *
 * Wird von `upgradeAudienceFieldToNote()` (existierendes Audience-Feld) und
 * `upgradeOrganizerFieldsToNote()` (Organizer + OrganizerEmail) genutzt. Die alte
 * `upgradeAudienceFieldToNote()`-Implementierung ist aus Kompatibilitätsgründen
 * unberührt geblieben — neue Migrationen sollten diesen Helper nutzen.
 */
export async function _upgradeTextFieldToNote(svc: EventService, listName: string, fieldName: string): Promise<void> {
  const tag = `_upgradeTextFieldToNote(${listName}.${fieldName})`;
  try {
    const fieldResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('${fieldName}')?$select=TypeAsString,FieldTypeKind`,
      SPHttpClient.configurations.v1
    );
    if (!fieldResp.ok) return;
    const fieldData = await fieldResp.json();
    const typeAsString: string = fieldData.TypeAsString || fieldData.d?.TypeAsString || '';
    const fieldTypeKind: number = fieldData.FieldTypeKind ?? fieldData.d?.FieldTypeKind ?? 0;
    if (typeAsString === 'Note' || fieldTypeKind === 3) return;
    if (typeAsString !== 'Text' && fieldTypeKind !== 2) {
      console.warn(`[DEX] ${tag}: unerwarteter Typ '${typeAsString}' (kind=${fieldTypeKind}) — skip.`);
      return;
    }

    const itemsResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,${fieldName}&$top=2000`,
      SPHttpClient.configurations.v1
    );
    if (!itemsResp.ok) return;
    const itemsData = await itemsResp.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const items: Array<any> = itemsData.value || itemsData.d?.results || [];
    const backup: Record<number, string> = {};
    for (const it of items) {
      const v = it[fieldName];
      if (v) backup[it.Id] = v;
    }
    console.warn(`[DEX] ${tag}: Backup ${Object.keys(backup).length} von ${items.length} Werten.`);

    try {
      await svc._post(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('${fieldName}')/deleteObject`,
        {}
      );
    } catch (e) {
      console.warn(`[DEX] ${tag}: Delete alte Spalte fehlgeschlagen, Migration abgebrochen:`, e);
      return;
    }

    try {
      await svc._post(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
        {
          '__metadata': { 'type': 'SP.FieldMultiLineText' },
          'Title': fieldName,
          'FieldTypeKind': 3,
          'Required': false,
          'RichText': false,
          'NumberOfLines': 4,
        }
      );
    } catch (e) {
      console.error(`[DEX] ${tag}: Konnte neue Note-Spalte nicht anlegen — Daten könnten verloren gehen:`, e, backup);
      return;
    }

    let restored = 0;
    let failed = 0;
    for (const idStr of Object.keys(backup)) {
      const id = Number(idStr);
      try {
        const resp = await svc._merge(
          `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`,
          { [fieldName]: backup[id] }
        );
        if (resp.ok) restored += 1;
        else failed += 1;
      } catch { failed += 1; }
    }
    console.warn(`[DEX] ${tag}: Migration fertig — ${restored} Werte zurückgeschrieben, ${failed} Fehler.`);
  } catch (e) {
    console.warn(`[DEX] ${tag} Error:`, e);
  }
}

/**
 * Berechtigungen für DEX_Events setzen
 */
async function setEventsListPermissions(svc: EventService, listName: string): Promise<void> {
  try {
    await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
      {}
    );

    const ownersResponse = await svc._sp.get(
      `${svc.siteUrl}/_api/web/associatedownergroup?$select=Id`,
      SPHttpClient.configurations.v1
    );
    if (ownersResponse.ok) {
      const ownersData = await ownersResponse.json();
      await svc._post(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${ownersData.Id}, roledefid=1073741829)`,
        {}
      );
    }

    const visitorsId = await svc.getVisitorsGroupId();
    if (visitorsId) {
      await svc._post(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741826)`,
        {}
      );
    }
  } catch {
    // Berechtigungen konnten nicht gesetzt werden
  }
}

/**
 * Default View einer Liste konfigurieren
 */
export async function configureDefaultView(svc: EventService, listName: string, fieldNames: string[], baseUrl?: string, opts?: { rebuild?: boolean }): Promise<void> {
  const url = baseUrl || svc.siteUrl;
  try {
    let existingFields: string[] = [];
    if (opts?.rebuild) {
      // Komplett neu aufbauen — SP-Defaults (Modified, Created, ID, Type,
      // Compliance-Tag, App Created By, ...) werden rausgeworfen.
      try {
        await svc._post(
          `${url}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields/removeallviewfields`,
          {}
        );
      } catch { /* ignore */ }
    }
    // IMMER die tatsächlich noch vorhandenen View-Felder lesen — auch im
    // Rebuild-Pfad. Hintergrund: SharePoints `addviewfield` ist NICHT
    // idempotent und antwortet mit HTTP 500, wenn das Feld bereits in der
    // View liegt. Greift `removeallviewfields` nicht (Permission/Throttle),
    // blieben die alten Felder drin → 500-Rauschen bei jedem Boot (z.B.
    // `addviewfield('DisableOutlook')`). Mit dem Re-Read überspringen wir
    // bereits vorhandene Felder und senden den fehlschlagenden Request gar
    // nicht erst ab.
    try {
      const existingResponse = await svc._sp.get(
        `${url}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields`,
        SPHttpClient.configurations.v1
      );
      if (existingResponse.ok) {
        const existingData = await existingResponse.json();
        if (existingData.Items) existingFields = existingData.Items;
        else if (existingData.d?.Items) existingFields = existingData.d.Items;
        else if (existingData.value) existingFields = existingData.value;
      }
    } catch { /* ignore — dann werden ggf. alle Felder versucht */ }

    for (const fieldName of fieldNames) {
      // Nur hinzufügen wenn noch nicht in der View
      if (existingFields.indexOf(fieldName) < 0) {
        await svc._post(
          `${url}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields/addviewfield('${fieldName}')`,
          {}
        );
      }
    }
  } catch {
    // View-Konfiguration ist optional
  }
}

/**
 * Column Formatting auf ein Feld setzen (z.B. Bild-Vorschau für URL-Spalten)
 */
async function setColumnFormatting(svc: EventService, listName: string, fieldName: string, formatJson: object): Promise<void> {
  try {
    const response = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$filter=InternalName eq '${fieldName}'&$select=Id`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) return;
    const data = await response.json();
    const field = data.value?.[0];
    if (!field) return;

    await svc._merge(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields('${field.Id}')`,
      { CustomFormatter: JSON.stringify(formatJson) }
    );
    // Column Formatting gesetzt
  } catch {
    // Column Formatting ist optional
  }
}
