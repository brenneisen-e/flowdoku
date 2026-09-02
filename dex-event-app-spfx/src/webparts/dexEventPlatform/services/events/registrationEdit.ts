/**
 * v30.66 — Modularisierung Stufe 2: Thema „Anmeldung lesen und ändern":
 * Wieder-Anmeldung einer abgemeldeten Zeile, Formularantworten und
 * Stammdaten ändern, Team zuordnen, Einwilligung/Einladung, Startnummer —
 * und das Lesen aller Zeilen einer Teilnehmerliste.
 *
 * `getAllRegistrations` WIRFT NICHT: Bei einem HTTP-Fehler bricht die
 * Schleife ab und liefert die bis dahin gelesenen Zeilen — bei 404 also `[]`.
 * Wer aus einem leeren Ergebnis auf „nicht vorhanden" schliesst, muss den
 * `onHttpError`-Rückruf nutzen (siehe CLAUDE.md, v29.3/v30.37).
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
// EventService als WERT-Import: nur für den deferred Zugriff auf die statische
// stripNoteWrapper — der Zyklus ist unkritisch, weil der Zugriff erst zur
// Laufzeit in Funktionskörpern passiert (wie REG_LIST_NAME).
import { EventService } from '../EventService';
import type { SPRegistration } from '../EventService';
import { REG_LIST_NAME } from '../EventService';

/**
 * Bestehende abgemeldete Registrierung reaktivieren.
 * Setzt Status zurück auf Angemeldet/Warteliste, löscht CancellationDate,
 * aktualisiert RegistrationDate und CustomData.
 */
export async function reactivateRegistration(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  firstName: string,
  surname: string,
  customData: Record<string, string>,
  status: string = 'Angemeldet',
  customFieldMap?: Record<string, string>,
  registeredByName?: string, // Audit: Name des Users der die Re-Anmeldung auslöst
  registeredByEmail?: string, // Audit: E-Mail des Users der die Re-Anmeldung auslöst
  proxyConsent?: string, // v18.74: Zustimmungs-Nachweis bei stellvertretender Re-Anmeldung
  starterType?: string, // v30.67: Gruppe bei geteilten Kapazitäten (wie registerForEvent)
  preferredStarterType?: string // v30.67: gewünschte Gruppe (Warteliste: nur diese, StarterType leer)
): Promise<boolean> {
  try {
    // ---- Permission-Checks (v3.9.2 / v3.9.3) ----
    // Lade die ParticipantEmail aus dem zu reaktivierenden Item und prüfe,
    // ob der aktuelle User dafür berechtigt ist. Plus Deadline-Check.
    // v7.30: Wir laden hier zusätzlich die existierende TeilnehmerID, damit
    // beim Reaktivieren die alte ID erhalten bleibt — Counter wird NUR
    // dann angefasst, wenn die alte ID null/0 ist (Legacy-Edge).
    let existingTeilnehmerId = 0;
    // v20.5: Teilnehmer-E-Mail in den aeusseren Scope heben, damit der
    // Autor-Set am Ende (stellvertretende Re-Anmeldung) sie nutzen kann.
    let reactivateParticipantEmail = '';
    try {
      const itemResp = await svc._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items(${itemId})?$select=ParticipantEmail,TeilnehmerID`,
        SPHttpClient.configurations.v1
      );
      const sessionEmail = (svc.context.pageContext.user.email || '').toLowerCase();
      let targetEmail = '';
      if (itemResp.ok) {
        const itemData = await itemResp.json();
        targetEmail = (itemData.ParticipantEmail || itemData.d?.ParticipantEmail || '').toLowerCase();
        reactivateParticipantEmail = targetEmail;
        const tnId = itemData.TeilnehmerID ?? itemData.d?.TeilnehmerID;
        if (typeof tnId === 'number' && tnId > 0) existingTeilnehmerId = tnId;
      }

      // Check A: für andere Person registrieren?
      if (targetEmail && targetEmail !== sessionEmail) {
        const allowed = await svc.canRegisterForOthers(subsiteUrl, targetEmail);
        if (!allowed) {
          console.warn(`[DEX] reactivateRegistration DENIED: ${sessionEmail} versuchte ${targetEmail} zu reaktivieren — nicht berechtigt.`);
          return false;
        }
      }

      // Check B: Deadline-Check (Event über SubsiteUrl finden)
      const subsiteEsc = encodeURIComponent(subsiteUrl.replace(/'/g, "''"));
      const evResp = await svc._sp.get(
        `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=SubsiteUrl eq '${subsiteEsc}'&$top=1&$select=RegistrationDeadline,OrganizerEmail`,
        SPHttpClient.configurations.v1
      );
      if (evResp.ok) {
        const evData = await evResp.json();
        const items = evData.value || evData.d?.results || [];
        if (items.length > 0) {
          const deadline = items[0].RegistrationDeadline || '';
          const orgStr: string = EventService.stripNoteWrapper(items[0].OrganizerEmail);
          const orgEmails = orgStr.split(';').map(s => s.trim().toLowerCase()).filter(Boolean);
          if (deadline) {
            const deadlineDate = new Date(deadline);
            if (!isNaN(deadlineDate.getTime()) && deadlineDate < new Date()) {
              const isEventOrganizer = orgEmails.indexOf(sessionEmail) >= 0;
              let isAdmin = false;
              try {
                const esc = sessionEmail.replace(/'/g, "''");
                const roleResp = await svc._sp.get(
                  `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Title eq '${encodeURIComponent(esc)}'&$top=1&$select=Role`,
                  SPHttpClient.configurations.v1
                );
                if (roleResp.ok) {
                  const rd = await roleResp.json();
                  const rItems = rd.value || rd.d?.results || [];
                  if (rItems.length > 0 && (rItems[0].Role === 'Admin' || rItems[0].Role === 'IT-Admin')) isAdmin = true;
                }
              } catch { /* ignore */ }
              if (!isEventOrganizer && !isAdmin) {
                console.warn(`[DEX] reactivateRegistration DENIED (deadline): ${sessionEmail} versuchte nach Deadline ${deadline} zu reaktivieren.`);
                return false;
              }
            }
          }
        }
      }
    } catch { /* bei Load-Fehler konservativ: weitermachen */ }
    // ---- Ende Permission-Checks ----

    // Reaktivierung = funktional eine Neuanmeldung mit existierendem Listen-
    // Item. Deshalb wird hier — analog zu registerForEvent — atomar eine
    // neue TeilnehmerID am Counter gezogen. Wer mal #12 war und reaktiviert,
    // bekommt jetzt z.B. die #87, also die nächst-freie ID am Ende der
    // Liste — exakt wie ein Neuzugang. Ohne diesen Schritt blieb der
    // Eintrag mit TeilnehmerID=null hängen, weil im Reaktivierungs-Pfad
    // niemand den DEX_IDReorder-Flow triggert.
    // v9.10: race-anfälliger max+1-Fallback entfernt (siehe Kommentar in
    // registerForEvent). Bei Counter-Outage bleibt TeilnehmerID undefined.
    void existingTeilnehmerId;
    const nextId = await svc.getNextTeilnehmerId(subsiteUrl);

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: Record<string, any> = {
      'Vorname': firstName,
      'Nachname': surname,
      'ParticipantName': `${firstName} ${surname}`,
      'Status': status,
      // v9.10: TeilnehmerID nur setzen wenn Counter sie geliefert hat.
      ...(typeof nextId === 'number' ? { 'TeilnehmerID': nextId } : {}),
      'RegistrationDate': new Date().toISOString(),
      'CancellationDate': null,
      'CustomData': JSON.stringify(customData),
    };

    // Audit: wer hat die Re-Anmeldung ausgelöst? (überschreibt den Wert von
    // der ursprünglichen Anmeldung, weil das faktisch eine neue Anmeldung ist)
    const auditName = registeredByName || svc.context.pageContext.user.displayName || '';
    const auditEmail = (registeredByEmail || svc.context.pageContext.user.email || '').toLowerCase();
    if (auditName) body['RegisteredByName'] = auditName;
    if (auditEmail) body['RegisteredByEmail'] = auditEmail;
    // v18.74: Zustimmungs-Nachweis bei stellvertretender Re-Anmeldung.
    if (proxyConsent) body['ProxyConsent'] = proxyConsent;
    // v30.67: Gruppe bei geteilten Kapazitäten mitschreiben — dasselbe Muster
    // wie registerForEvent, bei Warteliste wie switchSplitGroup (StarterType
    // leer, nur PreferredStarterType). Bisher nahm dieser zweite, ältere
    // Anmeldepfad die Gruppe gar nicht entgegen: Die Zeile trug weiter den
    // StarterType der ERSTEN Anmeldung, während reserveSeat schon einen Platz
    // in der neu gewählten Gruppe verbraucht hatte — Person in Gruppe A
    // gezählt, Platz in Gruppe B belegt, der nächste Zähler-Sync überbucht A.
    if (starterType) body['StarterType'] = starterType;
    if (preferredStarterType) {
      body['PreferredStarterType'] = preferredStarterType;
      if (status === 'Warteliste') body['StarterType'] = '';
    }

    if (customFieldMap) {
      for (const cfId of Object.keys(customData)) {
        if (cfId === 'salutation') continue;
        if (!customData[cfId]) continue;
        const spFieldName = customFieldMap[cfId];
        if (spFieldName) {
          body[spFieldName] = customData[cfId];
        } else {
          console.warn(`[DEX] Custom-Field '${cfId}' hat keine SharePoint-Spalte (spInternalName fehlt) — Wert wird nicht in Teilnehmerliste geschrieben. Bitte im Admin Center 'Custom Fields prüfen' ausführen.`);
        }
      }
    }

    let response = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      body
    );
    // v19.9: Wie bei registerForEvent — falls die ProxyConsent-Spalte auf
    // diesem (älteren) Event noch fehlt, scheitert der MERGE mit HTTP 400.
    // Einmal ohne das optionale Audit-Feld wiederholen, damit die
    // Re-Aktivierung nicht an einer fehlenden Spalte scheitert.
    if (!response.ok && body['ProxyConsent']) {
      console.warn('[DEX] reactivateRegistration: MERGE fehlgeschlagen — Retry OHNE ProxyConsent (Spalte evtl. nicht vorhanden). Bitte im Admin Center "Spalten fixen" ausführen.');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const retryBody: Record<string, any> = { ...body };
      delete retryBody['ProxyConsent'];
      response = await svc._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        retryBody
      );
    }
    if (!response.ok) return false;

    // v9.10: Post-Update Safety Net (siehe registerForEvent). Bei
    // Counter-Edge-Cases könnte der nächste Wert kollidieren — der
    // aeltere Eintrag (kleinere SP-Id) gewinnt, der spätere bekommt
    // fresh ID.
    if (typeof nextId === 'number' && nextId > 0) {
      try {
        const dupResp = await svc._sp.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,TeilnehmerID&$filter=TeilnehmerID eq ${nextId}&$top=10`,
          SPHttpClient.configurations.v1
        );
        if (dupResp.ok) {
          const dupData = await dupResp.json();
          const dupItems: Array<{ Id: number; TeilnehmerID: number }> = dupData.value || dupData.d?.results || [];
          if (dupItems.length > 1) {
            const minId = Math.min(...dupItems.map(d => d.Id));
            if (itemId !== minId) {
              const fresh = await svc.getNextTeilnehmerId(subsiteUrl);
              if (typeof fresh === 'number' && fresh > 0) {
                await svc._merge(
                  `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
                  { 'TeilnehmerID': fresh }
                );
                console.warn(`[DEX] Post-update dedup (reactivate): TeilnehmerID ${nextId} kollidierte, Item ${itemId} hat jetzt #${fresh}.`);
              }
            }
          }
        }
      } catch (err) {
        console.warn('[DEX] Post-update dedup check (reactivate) fehlgeschlagen:', err);
      }
    }

    // v20.5: Stellvertretende Re-Anmeldung → Teilnehmer zum Autor der Zeile
    // machen (analog registerForEvent), damit er sie in "Meine Events" sieht
    // und sich selbst abmelden kann. Best-effort (nur mit "Listen verwalten").
    if (reactivateParticipantEmail && auditEmail && reactivateParticipantEmail !== auditEmail) {
      await svc.trySetItemAuthor(subsiteUrl, REG_LIST_NAME, itemId, reactivateParticipantEmail);
    }

    return true;
  } catch {
    return false;
  }
}

/**
 * Schlanker MERGE-Helper auf ein einzelnes Teilnehmerlisten-Item — baut
 * keine ChangeLog-Logik, keine FieldMap-Auflösung, keine Default-Felder ein.
 * Genutzt für One-Shot-Migrationen (z.B. T-Shirt-Größen-Import), die direkt
 * bestimmte Felder (inkl. CustomData-JSON + einzelne SP-Spalten) setzen wollen.
 */
export async function mergeRegistrationFields(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  body: Record<string, unknown>
): Promise<boolean> {
  try {
    const response = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      body
    );
    return response.ok;
  } catch {
    return false;
  }
}

/**
 * Custom Data einer Registrierung aktualisieren (Teilnehmer ändert eigene Angaben).
 */
/**
 * v17.2: Bestehende Teilnehmer-Registrierung einem Team zuordnen
 * (PATCH der TeamId/TeamName/TeamLead-Felder auf einem schon existierenden
 * Item). Wird vom Admin-Center-Team-Management genutzt, wenn der
 * Organizer einen schon Angemeldeten ohne Team einem (neuen) Team
 * zuweist — vermeidet doppelte Anmeldung + Mail/Outlook-Spam.
 */
export async function assignRegistrationToTeam(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  teamId: string,
  teamName: string | undefined,
  isLead: boolean,
): Promise<boolean> {
  try {
    const body: Record<string, unknown> = {
      TeamId: teamId,
      TeamName: teamName || '',
      TeamLead: !!isLead,
    };
    const resp = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      body
    );
    return resp.ok;
  } catch (err) {
    console.warn('[DEX] assignRegistrationToTeam failed:', err);
    return false;
  }
}

export async function updateRegistrationData(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  customData: Record<string, string>,
  customFieldMap?: Record<string, string>,
  oldCustomData?: Record<string, string>,
  fieldLabelMap?: Record<string, string> // cf.id -> label
): Promise<boolean> {
  try {
    // Änderungen ermitteln
    const changes: string[] = [];
    const now = new Date().toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    if (oldCustomData && fieldLabelMap) {
      for (const key of Object.keys(customData)) {
        if (key === 'salutation') continue;
        const label = fieldLabelMap[key] || key;
        const oldVal = oldCustomData[key] || '';
        const newVal = customData[key] || '';
        if (oldVal !== newVal) {
          changes.push(`${label}: "${oldVal}" → "${newVal}"`);
        }
      }
    }
    const changeEntry = changes.length > 0 ? `[${now}] ${changes.join(', ')}` : '';

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: Record<string, any> = {
      'CustomData': JSON.stringify(customData),
      'LastModifiedDate': new Date().toISOString(),
    };

    // ChangeLog anhängen (bestehenden Log behalten)
    if (changeEntry) {
      try {
        const existing = await svc._sp.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})?$select=ChangeLog`,
          SPHttpClient.configurations.v1
        );
        if (existing.ok) {
          const data = await existing.json();
          const oldLog = data.ChangeLog || '';
          body['ChangeLog'] = oldLog ? `${changeEntry}\n${oldLog}` : changeEntry;
        }
      } catch {
        body['ChangeLog'] = changeEntry;
      }
    }

    if (customFieldMap) {
      for (const cfId of Object.keys(customData)) {
        if (cfId === 'salutation') continue;
        if (!customData[cfId]) continue;
        const spFieldName = customFieldMap[cfId];
        if (spFieldName) {
          body[spFieldName] = customData[cfId];
        } else {
          console.warn(`[DEX] Custom-Field '${cfId}' hat keine SharePoint-Spalte (spInternalName fehlt) — Wert wird nicht in Teilnehmerliste geschrieben. Bitte im Admin Center 'Custom Fields prüfen' ausführen.`);
        }
      }
    }

    const response = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      body
    );
    return response.ok;
  } catch {
    return false;
  }
}

/**
 * v8.0: Admin/Organizer kann Teilnehmerdaten direkt aus dem Admin Center
 * editieren. Erlaubt das Ändern von Anrede/Vorname/Nachname/Email/Phone/
 * Department/Location/JobTitle/Status sowie aller Custom-Felder. Schreibt
 * automatisch ChangeLog-Eintrag mit Wer/Wann/Was-Diff und setzt
 * LastModifiedDate.
 *
 * patch: Nur die echten Spalten-Werte (keine __metadata nötig — _merge
 * sendet odata=nometadata).
 * actor: Audit-Info des aufrufenden Users.
 * oldValues: zum Diff-Bauen, nur Felder mit oldValues[key] !== patch[key]
 * landen im ChangeLog.
 * fieldLabelMap: optional, mappt internal column name -> display label.
 */
export async function adminUpdateRegistration(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  patch: Record<string, unknown>,
  actor: { name: string; email: string },
  oldValues?: Record<string, unknown>,
  fieldLabelMap?: Record<string, string>
): Promise<boolean> {
  try {
    const changes: string[] = [];
    const now = new Date().toLocaleString('de-DE', {
      day: '2-digit', month: '2-digit', year: 'numeric',
      hour: '2-digit', minute: '2-digit',
    });
    if (oldValues) {
      for (const key of Object.keys(patch)) {
        const oldV = oldValues[key];
        const newV = patch[key];
        // Vergleich als String, damit number vs string nicht stört
        const oldStr = oldV === null || oldV === undefined ? '' : String(oldV);
        const newStr = newV === null || newV === undefined ? '' : String(newV);
        if (oldStr !== newStr) {
          const label = (fieldLabelMap && fieldLabelMap[key]) || key;
          changes.push(`${label}: "${oldStr}" → "${newStr}"`);
        }
      }
    }
    const changeEntry = changes.length > 0
      ? `[${now}] ${actor.name || actor.email}: ${changes.join(', ')}`
      : '';

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: Record<string, any> = { ...patch, 'LastModifiedDate': new Date().toISOString() };

    // ChangeLog anhängen (bestehenden Log behalten, neuestes oben)
    if (changeEntry) {
      try {
        const existing = await svc._sp.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})?$select=ChangeLog`,
          SPHttpClient.configurations.v1
        );
        if (existing.ok) {
          const data = await existing.json();
          const oldLog = data.ChangeLog || data.d?.ChangeLog || '';
          body['ChangeLog'] = oldLog ? `${changeEntry}\n${oldLog}` : changeEntry;
        } else {
          body['ChangeLog'] = changeEntry;
        }
      } catch {
        body['ChangeLog'] = changeEntry;
      }
    }

    const response = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      body
    );
    return response.ok;
  } catch (err) {
    console.warn('[DEX] adminUpdateRegistration error:', err);
    return false;
  }
}

/**
 * Eigene Registrierung für ein Event laden
 *
 * @param onHttpError v30.67: Meldet, dass die Abfrage gescheitert ist (HTTP-
 *   Status, 0 bei Netzwerkfehler). Ohne den Rückruf ist `null` zweideutig —
 *   „keine Zeile" und „konnte nicht lesen" sehen gleich aus. Genau daran
 *   ging der v30.14-Fail-closed vorbei: `_sp.get` wirft bei 429 nicht, das
 *   `.catch` im Aufrufer feuerte nie, und die Doppel-Anmelde-Prüfung sagte
 *   „nicht angemeldet". Wer aus `null` auf „nicht vorhanden" schließt, muss
 *   den Rückruf nutzen (dasselbe Muster wie `getAllRegistrations`).
 */
export async function getMyRegistration(
  svc: EventService,
  subsiteUrl: string,
  email: string,
  onHttpError?: (_status: number) => void
): Promise<SPRegistration | null> {
  try {
    // v27.11: $orderby=Id desc — bei mehreren Zeilen derselben Person
    // (Alt-Duplikate) IMMER die neueste nehmen. Vorher konnte $top=1 ohne
    // Sortierung eine alte 'Abgemeldet'-Zeile erwischen und den
    // Reaktivierungs-Pfad statt des Duplikat-Blocks auslösen.
    const response = await svc._sp.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=ParticipantEmail eq '${email.trim().replace(/'/g, "''")}'&$orderby=Id desc&$top=1`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) { if (onHttpError) onHttpError(response.status); return null; }
    const data = await response.json();
    if (data.value && data.value.length > 0) return data.value[0];
    return null;
  } catch {
    if (onHttpError) onHttpError(0);
    return null;
  }
}

/**
 * v24.36: Alle Anmeldungen einer Teilnehmerliste, die der eingeloggte User
 * STELLVERTRETEND für eine andere Person durchgeführt hat (Assistenz-Funktion).
 * Filter: `RegisteredByEmail = Akteur` UND `ParticipantEmail ≠ Akteur`.
 *
 * Hinweis zur Item-Level-Security: Auf den Teilnehmerlisten greift
 * ReadSecurity=2 („nur eigene Elemente"). Eine Assistenz (Contribute-User)
 * sieht daher nur die Zeilen, deren SharePoint-Autor sie selbst ist. Nach dem
 * v20.5-Autor-Wechsel (AuthorId → Teilnehmer) verliert die Assistenz den
 * Lesezugriff auf genau diese Zeilen — solange der `DEX_AccessFix`-Autor-Flow
 * NICHT eingerichtet ist (häufigster Fall), bleibt die Assistenz Autor und
 * sieht ihre Fremd-Anmeldungen weiterhin. Admin/Organizer eigener Events
 * sehen ohnehin alle Zeilen. Best-effort: liefert nur, was der Server
 * tatsächlich zurückgibt.
 */
export async function getProxyRegistrationsByActor(
  svc: EventService,
  subsiteUrl: string,
  actorEmail: string
): Promise<SPRegistration[]> {
  const me = (actorEmail || '').toLowerCase().trim();
  if (!me) return [];
  const out: SPRegistration[] = [];
  const esc = me.replace(/'/g, "''");
  let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=RegisteredByEmail eq '${esc}'&$orderby=Id asc&$top=5000`;
  while (url) {
    try {
      const response = await svc._sp.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) break;
      const data = await response.json();
      const page: SPRegistration[] = data.value || data.d?.results || [];
      for (const r of page) {
        const pe = (r.ParticipantEmail || '').toLowerCase().trim();
        if (pe && pe !== me) out.push(r);
      }
      url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
    } catch {
      break;
    }
  }
  return out;
}

/** v26.47: Externe stellvertretende Anmeldung als „Datenschutz-Rückmeldung
 *  offen" markieren (ConsentReview='Pending'). Lookup per E-Mail auf der
 *  jüngsten aktiven Zeile — läuft direkt nach der Registrierung. */
export async function markConsentPendingByEmail(svc: EventService, subsiteUrl: string, participantEmail: string): Promise<boolean> {
  try {
    const emailLc = (participantEmail || '').trim().toLowerCase();
    if (!emailLc) return false;
    const resp = await svc._sp.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,ParticipantEmail,Status&$orderby=Id desc&$top=200`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return false;
    const data = await resp.json();
    const items = (data.value || data.d?.results || []) as Array<{ Id: number; ParticipantEmail?: string; Status?: string }>;
    const hit = items.find(it => (it.ParticipantEmail || '').trim().toLowerCase() === emailLc && it.Status !== 'Abgemeldet');
    if (!hit) return false;
    const m = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${hit.Id})`,
      { 'ConsentReview': 'Pending' }
    );
    return m.ok;
  } catch (err) { console.warn('[DEX] markConsentPendingByEmail failed:', err); return false; }
}

/** v26.73: Fertigen .eml-Einladungs-Entwurf als Attachment an der Teilnehmer-
 *  Zeile ablegen (Zeile per E-Mail gefunden). Rückgabe = Item-Id (0 =
 *  fehlgeschlagen). Der Deeplink in der externen Instruktions-Mail holt genau
 *  diese Datei wieder — der Anhang darf per Deloitte-Mail-Regel nicht direkt
 *  mitgeschickt werden. Fester Dateiname `dxinvite--Einladung.eml`. */
export async function storeInviteEmlByEmail(svc: EventService, subsiteUrl: string, participantEmail: string, emlContent: string): Promise<number> {
  try {
    const emailLc = (participantEmail || '').trim().toLowerCase();
    if (!emailLc || !emlContent) return 0;
    const resp = await svc._sp.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,ParticipantEmail,Status&$orderby=Id desc&$top=500`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return 0;
    const data = await resp.json();
    const items = (data.value || data.d?.results || []) as Array<{ Id: number; ParticipantEmail?: string; Status?: string }>;
    const hit = items.find(it => (it.ParticipantEmail || '').trim().toLowerCase() === emailLc && it.Status !== 'Abgemeldet');
    if (!hit) return 0;
    const fileName = 'dxinvite--Einladung.eml';
    // Vorherige Version best-effort löschen (sonst 409 bei erneuter Anmeldung).
    try {
      await svc._sp.post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${hit.Id})/AttachmentFiles/getByFileName('${encodeURIComponent(fileName)}')`,
        SPHttpClient.configurations.v1,
        { headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE', 'Accept': 'application/json;odata=nometadata' } }
      );
    } catch { /* gab es noch nicht */ }
    const buf = new TextEncoder().encode(emlContent);
    const add = await svc._sp.post(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${hit.Id})/AttachmentFiles/add(FileName='${encodeURIComponent(fileName)}')`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' }, body: buf.buffer as ArrayBuffer }
    );
    return add.ok ? hit.Id : 0;
  } catch (err) { console.warn('[DEX] storeInviteEmlByEmail failed:', err); return 0; }
}

/** v26.73: Den an der Teilnehmer-Zeile abgelegten .eml-Entwurf (per Item-Id)
 *  wieder auslesen — für den Download-Deeplink. */
export async function getInviteEmlByItem(svc: EventService, subsiteUrl: string, itemId: number): Promise<{ fileName: string; content: string } | null> {
  try {
    if (!subsiteUrl || !itemId) return null;
    const resp = await svc._sp.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${Number(itemId)})/AttachmentFiles/getByFileName('${encodeURIComponent('dxinvite--Einladung.eml')}')/$value`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } }
    );
    if (!resp.ok) return null;
    const content = await resp.text();
    if (!content) return null;
    return { fileName: 'Einladung.eml', content };
  } catch (err) { console.warn('[DEX] getInviteEmlByItem failed:', err); return null; }
}

/** v26.47: Datenschutz-Rückmeldung der externen Person bestätigen —
 *  ConsentReview zurücksetzen (Button in der Teilnehmerliste). */
export async function confirmConsentReview(svc: EventService, subsiteUrl: string, itemId: number, meta?: { eventId?: string; eventTitle?: string; participantName?: string }): Promise<boolean> {
  try {
    const m = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      { 'ConsentReview': '' }
    );
    if (m.ok && meta?.eventId) {
      try {
        await svc.writeChangeLog({
          action: 'ExternalConsentConfirmed', targetType: 'Participant', targetId: String(itemId),
          targetName: meta.participantName || '', eventId: meta.eventId, eventTitle: meta.eventTitle || '',
          details: { note: 'Datenschutz-Rückmeldung der externen Person bestätigt (v26.47).' },
        });
      } catch { /* Audit best-effort */ }
    }
    return m.ok;
  } catch (err) { console.warn('[DEX] confirmConsentReview failed:', err); return false; }
}

/**
 * v30.48: Spalte `Startnummer` in der Teilnehmerliste sicherstellen.
 *
 * Eigene Spalte statt Custom-Field: Die Startnummer soll im Export, in der
 * Teilnehmerliste und in der Check-in-Suche auftauchen — ein Custom-Field
 * wäre pro Event konfigurierbar, aber nicht durchsuchbar. Text und nicht
 * Zahl, weil Veranstalter-Nummern führende Nullen und Präfixe haben können;
 * gerechnet wird damit ohnehin nie.
 *
 * Idempotent: Existiert die Spalte, passiert nichts. Wirft nicht — der
 * Import meldet den Fehlschlag selbst, statt hier abzubrechen.
 */
export async function ensureStartNumberColumn(svc: EventService, subsiteUrl: string): Promise<boolean> {
  try {
    const resp = await svc._sp.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields?$filter=InternalName eq 'Startnummer'&$select=InternalName`,
      SPHttpClient.configurations.v1
    );
    if (resp.ok) {
      const d = await resp.json();
      const arr = d.value || d.d?.results || [];
      if (arr.length > 0) return true;
    }
    await svc._post(`${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`, {
      '__metadata': { 'type': 'SP.Field' },
      'Title': 'Startnummer',
      'FieldTypeKind': 2, // Text
      'Required': false,
    });
    return true;
  } catch (e) {
    console.warn('[DEX] ensureStartNumberColumn failed:', e);
    return false;
  }
}

/**
 * Alle Registrierungen für ein Event laden (nur für Organizer/Admin)
 */
/**
 * @param onHttpError v29.3: Meldet, dass der Lesevorgang abgebrochen wurde,
 *   statt ihn — wie bisher — still zu verschlucken. Ohne diesen Rückruf ist
 *   „Liste existiert nicht" (404, z.B. nach dem 3-Monats-Löschkonzept) von
 *   „Liste ist leer" nicht zu unterscheiden: Beides kam als `[]` zurück.
 *   Genau daran hat sich `analyzeRegistryAgainstLists` verschluckt.
 *   `status` ist der HTTP-Status, `0` bei einem Netz-/Parse-Fehler.
 */
export async function getAllRegistrations(svc: EventService, subsiteUrl: string, onHttpError?: (_status: number) => void): Promise<SPRegistration[]> {
  const allItems: SPRegistration[] = [];
  // $top=5000 ist das SP-REST-Maximum pro Page. Damit fallen Events bis zu
  // 5000 Teilnehmern in einen einzigen Response — keine Pagination-Edgecases
  // mit fehlendem nextLink. Bei größeren Listen folgen wir dem nextLink
  // weiter (Schleife unten). Vorher stand hier $top=500, was bei Events mit
  // ≥500 Teilnehmern zu fehlenden Einträgen führte: SharePoint liefert
  // bei $orderby+$top in Kombination mit Item-Level-Security nicht
  // zuverlässig nextLink, wenn die erste Page exakt voll ist.
  // v27.12: $select=*,Author/… + $expand=Author — der Zeilen-Autor dient als
  // Fallback für „Registriert von", wenn die Zeile nicht über die App
  // angelegt wurde (RegisteredBy* leer). '*' behält alle Skalar-Felder,
  // Verhalten ist sonst identisch zum bisherigen Query ohne $select.
  let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=*,Author/Title,Author/EMail&$expand=Author&$orderby=Id asc&$top=5000`;

  while (url) {
    try {
      const response = await svc._sp.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) { if (onHttpError) onHttpError(response.status); break; }
      const data = await response.json();
      // Beide OData-Formate abdecken: nometadata (data.value) UND verbose
      // (data.d.results). Vorher nur data.value — bei verbose-Response
      // wären null Items dazugekommen.
      const page = data.value || data.d?.results || [];
      allItems.push(...page);
      url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
    } catch {
      if (onHttpError) onHttpError(0);
      break;
    }
  }
  return allItems;
}

/**
 * v19.28: Eine Teilnehmer-Registrierung endgültig aus der Subsite-Liste
 * löschen (hartes DELETE, kein Recycle-Bin). Use-Case: abgemeldete
 * Test-Anmeldungen aus der Abmeldungen-Liste entfernen, damit die Übersicht
 * sauber bleibt. Die Berechtigung (Admin/Organizer) wird in der UI geprüft.
 */
export async function deleteRegistration(svc: EventService, subsiteUrl: string, itemId: number): Promise<boolean> {
  try {
    const resp = await svc._delete(`${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`);
    return resp.ok;
  } catch (err) {
    console.warn('[DEX] deleteRegistration failed:', err);
    return false;
  }
}
