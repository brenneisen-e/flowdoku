/**
 * v30.66 — Modularisierung Stufe 2: Thema „Teilnehmerliste prüfen und
 * reparieren". `diagnoseRegistrationList` sagt, welche Spalten einer
 * Bestands-Teilnehmerliste fehlen, `fixRegistrationListColumns` legt sie nach.
 *
 * Beides ist Bestandspflege: Neue Releases bringen neue Spalten mit, ältere
 * Event-Subsites haben sie nicht — und ein MERGE auf eine fehlende Spalte
 * endet in HTTP 400, nicht in einer Fehlermeldung in der App.
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService, CustomField } from '../EventService';
import { COUNTER_LIST_NAME, REG_LIST_NAME } from '../EventService';

/**
 * Spalten der Teilnehmerliste fixen: fehlende Spalten anlegen, View-Reihenfolge korrigieren.
 * Kann auf bestehenden Events ausgeführt werden um die Liste nachträglich zu aktualisieren.
 */
/**
 * v30.58: Prüft EINE Teilnehmerliste gegen die Abfragefelder ihres Events —
 * ohne etwas zu ändern.
 *
 * Der Anlass: Fehlt auf einer Liste die Spalte zu einem Abfragefeld, lehnt
 * SharePoint den GANZEN Insert ab (HTTP 400, „The field or property … does
 * not exist"). Betroffen sind dann nur die Personen, die dieses Feld
 * ausfüllen — deshalb sieht es aus wie ein zufälliger Einzelfall und nicht
 * wie ein Struktur-Fehler. Genau so ist die fehlende Klammer-Zeile
 * entstanden: Die übergreifenden Hauptevent-Felder gehen NUR auf die
 * Klammer-Liste, also scheitert auch nur dort der Schreibvorgang.
 *
 * Bewusst lesend: „Was ist kaputt?" und „repariere es" sind zwei Fragen.
 * Wer erst sehen will, was los ist, soll dafür nichts verändern müssen.
 */
export async function diagnoseRegistrationList(
  svc: EventService,
  subsiteUrl: string,
  customFields: Array<{ id: string; label: string; spInternalName?: string }>
): Promise<{ ok: boolean; listMissing: boolean; missingColumns: Array<{ id: string; label: string; column: string }>; error?: string }> {
  try {
    const resp = await svc._sp.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields?$filter=Hidden eq false&$top=500&$select=InternalName,Title`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) {
      // 404 heißt: Die Liste gibt es nicht (mehr). Das ist eine ANDERE
      // Aussage als „Spalte fehlt" und darf nicht damit vermischt werden —
      // dieselbe Unterscheidung wie bei getAllRegistrations (v29.3).
      return { ok: false, listMissing: resp.status === 404 || resp.status === 410, missingColumns: [], error: `HTTP ${resp.status}` };
    }
    const data = await resp.json();
    const rows: Array<{ InternalName?: string; Title?: string }> = data.value || data.d?.results || [];
    const have = new Set<string>();
    for (const r of rows) {
      if (r.InternalName) have.add(r.InternalName.toLowerCase());
      if (r.Title) have.add(r.Title.toLowerCase());
    }
    const missing: Array<{ id: string; label: string; column: string }> = [];
    for (const f of (customFields || [])) {
      const col = (f.spInternalName || '').trim();
      // Ohne spInternalName wurde das Feld noch nie auf eine Liste
      // geschrieben — dann ist nicht die Liste schuld, sondern die
      // Zuordnung fehlt. Beides meldet „Spalten fixen" als zu erledigen.
      if (!col) { missing.push({ id: f.id, label: f.label, column: '(noch nicht zugeordnet)' }); continue; }
      if (!have.has(col.toLowerCase())) missing.push({ id: f.id, label: f.label, column: col });
    }
    return { ok: true, listMissing: false, missingColumns: missing };
  } catch (e) {
    return { ok: false, listMissing: false, missingColumns: [], error: String((e as Error)?.message || e) };
  }
}

export async function fixRegistrationListColumns(
  svc: EventService,
  subsiteUrl: string,
  eventContext?: {
    isB2Run?: boolean;  // Event hat Durchstarter/Funstarter Kapazität
    hasQuiz?: boolean;  // Event hat Quizfragen
    customFields?: CustomField[]; // Event-spezifische Custom-Fields — fehlende SP-Spalten werden angelegt
  },
  // v11.56: Optionaler Confirm-Callback. Wird aufgerufen, wenn Duplikat-Spalten
  // erkannt wurden, BEVOR irgendetwas gelöscht wird. Liefert der Callback false,
  // werden die Duplikate übersprungen (die Hauptfix-Logik läuft trotzdem).
  confirmDeleteDuplicates?: (count: number, titles: string[]) => boolean | Promise<boolean>
): Promise<{ added: string[]; removed: string[]; viewFixed: boolean; customFieldMap?: Record<string, string>; duplicatesRemoved?: string[]; duplicatesWithData?: string[] }> {
  const added: string[] = [];
  const removed: string[] = [];
  const duplicatesRemoved: string[] = [];
  const duplicatesWithData: string[] = [];

  // Bestehende Felder laden — InternalName + Title beide nehmen, damit wir per Title
  // dedupen können (siehe v11.56: alte Builds haben durch fehlgeschlagene Existenz-
  // checks beim wiederholten "Spalten fixen" pro Custom-Field 50+ Duplikate angelegt).
  const fieldsResp = await svc._sp.get(
    `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields?$filter=Hidden eq false&$top=500&$select=InternalName,Title`,
    SPHttpClient.configurations.v1
  );
  const existingFieldsList: string[] = [];
  const existingByInternal: Set<string> = new Set();
  const existingByTitle: Map<string, Array<{ internalName: string }>> = new Map();
  if (fieldsResp.ok) {
    const fieldsData = await fieldsResp.json();
    const fields = fieldsData.value || fieldsData.d?.results || [];
    for (const f of fields) {
      const internalName: string = String(f.InternalName || '');
      const title: string = String(f.Title || '');
      if (!internalName) continue;
      existingFieldsList.push(internalName);
      existingByInternal.add(internalName);
      if (title) {
        const arr = existingByTitle.get(title) || [];
        arr.push({ internalName });
        existingByTitle.set(title, arr);
      }
    }
  }

  // ===== DEDUPE-PASS (v11.56) =====
  // Pro Title: wenn mehr als ein Feld diesen Titel hat → die uebrigen Felder löschen,
  // sofern sie leer sind (keine Items mit Wert in der Spalte). Erstes Feld bleibt
  // immer erhalten. Felder mit Daten werden gemeldet (duplicatesWithData) und nicht
  // automatisch gelöscht — der User soll sie manuell prüfen.
  const duplicateTitles: Array<{ title: string; entries: Array<{ internalName: string }> }> = [];
  existingByTitle.forEach((entries, title) => {
    if (entries.length > 1) {
      duplicateTitles.push({ title, entries });
    }
  });
  if (duplicateTitles.length > 0) {
    // Vor dem Löschen den Aufrufer fragen — Operation ist irreversibel.
    if (confirmDeleteDuplicates) {
      const count = duplicateTitles.reduce((sum, d) => sum + (d.entries.length - 1), 0);
      const titles = duplicateTitles.map(d => d.title);
      const ok = await Promise.resolve(confirmDeleteDuplicates(count, titles));
      if (!ok) {
        // Cleanup überspringen — nur den Hauptfix laufen lassen
        duplicateTitles.length = 0;
      }
    }
  }
  if (duplicateTitles.length > 0) {
    // Pro Duplikat-Set: den ersten Eintrag behalten, für alle weiteren prüfen ob leer.
    for (const dup of duplicateTitles) {
      // entries[0] bleibt erhalten
      for (let i = 1; i < dup.entries.length; i++) {
        const candidate = dup.entries[i];
        let isEmpty = false;
        // v11.67: SP truncated InternalNames auf 32 Zeichen. Wenn die Truncation
        // mitten in einer `_xXXXX_`-Encoding-Sequenz liegt (z.B.
        // `ADMIN_x0020__x002d__x0020_Who_x00` — die letzten Zeichen `_x00`
        // sind eine angeschnittene `_x0020_`-Sequenz), wirft SP HTTP 400 auf
        // jeden OData-`$filter`-Versuch. Solche Spalten werden hier nicht
        // geprüft → konservativ als „hat Daten" behandelt (kein Auto-
        // Löschen). Der Admin kann sie über die SP-Listen-UI manuell
        // entfernen, wenn sie wirklich leer sind.
        const looksTruncated = candidate.internalName.length === 32
          && /_x[0-9a-f]{1,3}$/i.test(candidate.internalName);
        if (looksTruncated) {
          if (duplicatesWithData.indexOf(dup.title) < 0) duplicatesWithData.push(dup.title);
          continue;
        }
        try {
          const checkUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=${candidate.internalName} ne null&$top=1&$select=ID`;
          const checkResp = await svc._sp.get(checkUrl, SPHttpClient.configurations.v1);
          if (checkResp.ok) {
            const data = await checkResp.json();
            const items = data.value || data.d?.results || [];
            isEmpty = items.length === 0;
          }
        } catch { isEmpty = false; }
        if (isEmpty) {
          try {
            const delResp = await svc._post(
              `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields/getbyinternalnameortitle('${candidate.internalName}')/deleteObject`,
              {}
            );
            if (delResp.ok) {
              duplicatesRemoved.push(candidate.internalName);
              // Aus existingByInternal + existingFieldsList rausziehen
              existingByInternal.delete(candidate.internalName);
              const idx = existingFieldsList.indexOf(candidate.internalName);
              if (idx >= 0) existingFieldsList.splice(idx, 1);
            }
          } catch { /* löschen fehlgeschlagen — weiter */ }
        } else {
          // Daten vorhanden — Title für manuelle Prüfung melden (nur einmal pro Title)
          if (duplicatesWithData.indexOf(dup.title) < 0) duplicatesWithData.push(dup.title);
        }
      }
      // existingByTitle entsprechend bereinigen: nur die nicht-gelöschten Einträge behalten
      const remaining = dup.entries.filter(e => existingByInternal.has(e.internalName));
      existingByTitle.set(dup.title, remaining);
    }
  }

  // Fehlende Basis-Spalten anlegen. StarterType/Quiz-Felder sind feature-spezifisch:
  // - StarterType/PreferredStarterType: nur für B2Run-Events mit Split-Kapazität
  // - QuizScore/QuizAnswers/QuizCompletedAt: nur für Events mit Quizfragen
  // Wird das Event ohne eventContext gefixt (kein Aufrufer-seitiger Flag), lassen wir
  // feature-spezifische Spalten raus, damit sie nicht unbegründet auf jedem
  // Teilnehmerlisten-Schema auftauchen.
  const requiredFields: Array<{ title: string; type: number; choices?: string[]; metaType?: string }> = [
    { title: 'Anrede', type: 6, choices: ['Frau', 'Herr', 'Divers'], metaType: 'SP.FieldChoice' },
    { title: 'Company', type: 2 },           // v24.29: Unternehmenszugehörigkeit / Rechtsträger (aus dem Profil)
    { title: 'RegisteredByName', type: 2 },  // Audit: Name des Users der die Anmeldung durchgeführt hat
    { title: 'RegisteredByEmail', type: 2 }, // Audit: E-Mail des Users der die Anmeldung durchgeführt hat
    { title: 'ProxyConsent', type: 3 },      // v18.74: Zustimmungs-Nachweis bei stellvertretender Anmeldung (Note)
    { title: 'CancelledByName', type: 2 },   // Audit: Name des Users der die Abmeldung ausgelöst hat
    { title: 'CancelledByEmail', type: 2 },  // Audit: E-Mail des Users der die Abmeldung ausgelöst hat
    { title: 'CheckedInDate', type: 4 },     // v7.16: Check-In-Audit — Zeitpunkt
    { title: 'CheckedInByName', type: 2 },   // v7.16: Check-In-Audit — Helfer-Name
    { title: 'CheckedInByEmail', type: 2 },  // v7.16: Check-In-Audit — Helfer-E-Mail
    // v19.3: Nachrück-Audit-Spalten auch beim „Spalten fixen" nachziehen, damit
    // der DEX_IDReorder-Flow (und der App-Button) sie auf Bestands-Events
    // beschreiben kann → „Nachgerückt am / Ersetzt / Ersetzt durch" in der App.
    { title: 'PromotedDate', type: 4 },              // DateTime — Zeitpunkt des Nachrückens
    { title: 'ReplacedParticipantEmail', type: 2 },  // E-Mail der Person, deren Cancel den Platz freigab
    { title: 'ReplacedByParticipantEmail', type: 2 },// E-Mail der nachrückenden Person (Spiegelbild)
    { title: 'OverbookReview', type: 2 },    // v11.36: Überbuchungs-Review-Marker
    { title: 'ConsentReview', type: 2 },     // v26.47: Externe Anmeldung — 'Pending' = Datenschutz-Rückmeldung offen
    { title: 'TeamId', type: 2 },            // v11.82: UUID einer Team-Anmeldung (leer = Solo)
    { title: 'TeamLead', type: 8 },          // v11.82: Boolean — true für die anmeldende Person
    { title: 'TeamName', type: 2 },          // v11.82: optionaler frei wählbarer Team-Name
  ];
  if (eventContext?.isB2Run) {
    requiredFields.push(
      { title: 'StarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' },
      { title: 'PreferredStarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' },
      // v10.13: b2run_leistungsnachweis ist ein virtuelles Feld das die
      // RegistrationPage hardcoded hinzufügt wenn durchstarterRequiresProof
      // aktiv ist — es ist NICHT Teil der regulären customFields, daher
      // muss die SP-Spalte hier explizit angelegt werden, sonst kippt die
      // Anmeldung mit HTTP 400 'Field not found'. Wird auf jedem B2Run-Event
      // angelegt (egal ob proof-flag aktuell aktiv ist) — die Spalte ist
      // klein und stört nicht wenn ungenutzt.
      { title: 'b2run_leistungsnachweis', type: 2 }
    );
  }
  if (eventContext?.hasQuiz) {
    requiredFields.push(
      { title: 'QuizScore', type: 9 },
      { title: 'QuizAnswers', type: 3 },
      { title: 'QuizCompletedAt', type: 4 }
    );
  }

  // Feature-spezifische Spalten, die auf diesem Event NICHT gebraucht werden,
  // aktiv löschen (z.B. StarterType auf einem Nicht-B2Run-Event). Das ist
  // irreversibel — eventuelle Daten in diesen Spalten gehen verloren. Ist aber
  // vom User explizit gewünscht, damit die Teilnehmerliste pro Event-Typ
  // sauber bleibt.
  const deletableFields: string[] = [];
  if (!eventContext?.isB2Run) {
    deletableFields.push('StarterType', 'PreferredStarterType');
  }
  if (!eventContext?.hasQuiz) {
    deletableFields.push('QuizScore', 'QuizAnswers', 'QuizCompletedAt');
  }
  for (const fieldName of deletableFields) {
    if (existingFieldsList.indexOf(fieldName) >= 0) {
      try {
        const delResp = await svc._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields/getbytitle('${fieldName}')/deleteObject`,
          {}
        );
        if (delResp.ok) {
          removed.push(fieldName);
          // aus existingFieldsList rausziehen damit die View-Logik weiter unten
          // den Feldnamen nicht mehr als "noch vorhanden" betrachtet.
          const idx = existingFieldsList.indexOf(fieldName);
          if (idx >= 0) existingFieldsList.splice(idx, 1);
        }
      } catch { /* Feld konnte nicht gelöscht werden - weitermachen */ }
    }
  }

  for (const f of requiredFields) {
    if (existingFieldsList.indexOf(f.title) < 0) {
      const payload: Record<string, unknown> = {
        '__metadata': { 'type': f.metaType || 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      };
      if (f.choices) {
        payload['Choices'] = { 'results': f.choices };
      }
      const resp = await svc._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`, payload
      );
      if (resp.ok) added.push(f.title);
    }
  }

  // Custom Fields pro Event: wenn spInternalName leer oder die Spalte fehlt,
  // jetzt anlegen. Der Aufrufer bekommt customFieldMap zurück und kann das
  // Event-Item persistieren (spInternalName für jede cf.id).
  const customFieldMap: Record<string, string> = {};
  if (eventContext?.customFields && eventContext.customFields.length > 0) {
    // Post-Fix Felder-Snapshot nach Basis-Anlage
    let currentFields = [...existingFieldsList, ...added];
    for (const cf of eventContext.customFields) {
      if (!cf.label || !cf.label.trim()) continue;
      // v19.0: Dokument-Felder bekommen keine Spalte (Datei = Attachment).
      if (cf.type === 'document') continue;
      // Wenn spInternalName schon gesetzt und Feld existiert: übernehmen.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const existingSp: string = String((cf as any).spInternalName || '');
      if (existingSp && currentFields.indexOf(existingSp) >= 0) {
        customFieldMap[cf.id] = existingSp;
        continue;
      }
      // v11.56: Wenn spInternalName fehlt oder nicht zur Liste passt, aber ein
      // Feld mit demselben Title bereits existiert: dieses InternalName übernehmen,
      // statt eine Duplikat-Spalte anzulegen. Das ist die Hauptursache der
      // 100x-Duplikate-Misere (P/D MEETING0, P/D MEETING1, ...).
      const titleMatches = existingByTitle.get(cf.label) || [];
      if (titleMatches.length > 0) {
        const firstInternal = titleMatches[0].internalName;
        customFieldMap[cf.id] = firstInternal;
        if (currentFields.indexOf(firstInternal) < 0) currentFields.push(firstInternal);
        continue;
      }
      // Feld-Payload je nach Typ
      let fieldPayload: Record<string, unknown>;
      if (cf.type === 'select' && cf.options && cf.options.length > 0) {
        fieldPayload = { '__metadata': { 'type': 'SP.FieldChoice' }, 'Title': cf.label, 'FieldTypeKind': 6, 'Required': false, 'Choices': { 'results': cf.options } };
      } else if (cf.type === 'number') {
        fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 9, 'Required': false };
      } else if (cf.type === 'checkbox') {
        fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 8, 'Required': false };
      } else if (cf.type === 'user') {
        // user-Picker wird als Text gespeichert ("Name <email>").
        fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 3, 'Required': false };
      } else {
        fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 2, 'Required': false };
      }
      try {
        const resp = await svc._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`, fieldPayload
        );
        if (resp.ok) {
          const createdField = await resp.json().catch(() => null);
          const internalName = createdField?.InternalName || createdField?.d?.InternalName || '';
          if (internalName) {
            customFieldMap[cf.id] = internalName;
            added.push(internalName);
            currentFields = currentFields.concat([internalName]);
            // Title-Map aktualisieren, damit ein zweites cf mit gleichem Label
            // im selben Durchlauf (z.B. zwei Custom-Fields mit identischem Title)
            // das gerade angelegte Feld wiederverwendet, statt erneut zu erzeugen.
            const arr = existingByTitle.get(cf.label) || [];
            arr.push({ internalName });
            existingByTitle.set(cf.label, arr);
          }
        }
      } catch { /* nächstes Feld */ }
    }
  }

  // Default View komplett neu aufbauen (Reihenfolge: TeilnehmerID, Anrede, Vorname, Nachname, ...)
  let viewFixed = false;
  try {
    // Alle bestehenden Felder aus der View entfernen
    await svc._post(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/defaultview/viewfields/removeallviewfields`,
      {}
    );

    // Felder in gewünschter Reihenfolge hinzufügen. StarterType/Quiz-Spalten
    // werden nur eingebaut, wenn sie tatsächlich auf der Liste existieren —
    // auf Nicht-B2Run- bzw. Nicht-Quiz-Events sollen sie nicht auftauchen.
    const viewFieldsCore = [
      'TeilnehmerID', 'Anrede', 'Vorname', 'Nachname', 'ParticipantEmail',
      'Department', 'Location', 'JobTitle', 'Company', 'Phone',
    ];
    const viewFields: string[] = [...viewFieldsCore];
    // Post-Fix-Feldliste (bestehende + gerade hinzugefügte) für die Existenz-Checks
    const postFixFields = existingFieldsList.concat(added);
    if (postFixFields.indexOf('StarterType') >= 0) viewFields.push('StarterType');
    if (postFixFields.indexOf('PreferredStarterType') >= 0) viewFields.push('PreferredStarterType');
    viewFields.push('Status', 'RegistrationDate');
    if (postFixFields.indexOf('RegisteredByName') >= 0) viewFields.push('RegisteredByName');
    if (postFixFields.indexOf('RegisteredByEmail') >= 0) viewFields.push('RegisteredByEmail');
    if (postFixFields.indexOf('ProxyConsent') >= 0) viewFields.push('ProxyConsent');
    viewFields.push('CancellationDate');

    // Wir blenden SP-System-Spalten komplett aus (Modified, Created, ID, Type,
    // Compliance Asset Id, Retention Label, etc.) — nur funktionelle Felder der
    // App + Custom Fields kommen in die View.
    const systemBlocklist = new Set([
      'ID', '_UIVersionString', 'Edit', 'LinkTitle', 'LinkTitleNoMenu',
      'LinkFilename', 'LinkFilenameNoMenu', 'DocIcon', 'FileLeafRef',
      'Modified', 'Created', 'Editor', 'Author', 'CreatedBy', 'ModifiedBy',
      'Title', 'ParticipantName',
      'ContentType', 'ContentTypeId', 'Attachments',
      'AppAuthor', 'AppEditor', 'App Created By', 'App Modified By',
      'Type', 'ItemChildCount', 'FolderChildCount',
      'ComplianceAssetId', '_ComplianceTag', '_ComplianceTagWrittenTime',
      '_ComplianceTagUserId', 'TaxCatchAll', 'TaxCatchAllLabel',
      'SMTotalFileStreamSize', 'SMTotalSize', 'SortBehavior',
      'OData__UIVersionString', 'OData__HasCopyDestinations',
      'LastModifiedDate', 'ChangeLog', 'CustomData',
      '_CopySource', 'owshiddenversion', 'WorkflowVersion', 'WorkflowInstanceID',
      'ItemIsRecord', '_HasEncryptedContent', '_IsRecord', '_IsRecordApplied',
      'InstanceID', 'Order', 'GUID', 'FileSizeDisplay', 'MetaInfo',
      'ParentUniqueId', 'AccessPolicy', 'HasUniqueRoleAssignments',
      'Restricted', 'Type0', 'ServerUrl', 'EncodedAbsUrl', 'BaseName',
      'FileType', 'HTML_x0020_File_x0020_Type', '_EditMenuTableStart',
      '_EditMenuTableStart2', '_EditMenuTableEnd', 'PermMask',
    ]);
    // Bereits zur View hinzugefügt — nicht doppelt anfassen
    const alreadyAdded = new Set(viewFields);
    // v11.82: Team-Spalten kommen ans Ende der View — nach allen
    // Custom-Fields, damit sie nicht zwischen den event-spezifischen
    // Antwortspalten landen. Hier merken und im Post-Loop überspringen.
    const teamTailFields = ['TeamId', 'TeamLead', 'TeamName'];
    const teamTailSet = new Set(teamTailFields);
    // Kompletter Feld-Stand NACH dem Fix (bestehende + neu angelegte),
    // damit neu angelegte Custom-Fields auch in die View kommen.
    for (const fn of postFixFields) {
      if (alreadyAdded.has(fn)) continue;
      if (systemBlocklist.has(fn)) continue;
      if (fn.charAt(0) === '_') continue;
      if (teamTailSet.has(fn)) continue; // ans Ende
      viewFields.push(fn);
      alreadyAdded.add(fn);
    }
    // Team-Spalten jetzt am Ende anhängen (nur die, die tatsächlich existieren).
    for (const fn of teamTailFields) {
      if (alreadyAdded.has(fn)) continue;
      if (postFixFields.indexOf(fn) < 0) continue;
      viewFields.push(fn);
      alreadyAdded.add(fn);
    }

    for (const fn of viewFields) {
      await svc._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/defaultview/viewfields/addviewfield('${fn}')`,
        {}
      );
    }
    viewFixed = true;
  } catch {
    console.warn('[DEX] View-Reihenfolge konnte nicht gesetzt werden');
  }

  // v7.28: Counter-Liste für atomare TeilnehmerID-Vergabe anlegen
  // (oder seeden mit dem aktuellen Max-Wert wenn schon vorhanden).
  try {
    const counterResult = await svc.ensureCounterList(subsiteUrl);
    if (counterResult.created) {
      added.push(`Counter-Liste ${COUNTER_LIST_NAME} (atomare TeilnehmerID-Vergabe, seeded mit ${counterResult.seededValue})`);
    } else if (counterResult.seededValue !== undefined) {
      added.push(`Counter-Item nachgeseedet (NextValue=${counterResult.seededValue})`);
    }
  } catch {
    console.warn('[DEX] Counter-Liste konnte nicht angelegt werden');
  }

  return { added, removed, viewFixed, customFieldMap: Object.keys(customFieldMap).length > 0 ? customFieldMap : undefined };
}
