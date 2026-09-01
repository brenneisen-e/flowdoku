/**
 * v30.66 — Modularisierung Stufe 2: Thema „Nachrücken von der Warteliste".
 *
 * Nachgerückt wird nur beim Abmelden — eine Kapazitätserhöhung erzeugt kein
 * Ereignis, an dem etwas hinge (dafür gibt es die Aktion „Freie Plätze mit
 * Warteliste füllen"). Bei geteilten Kapazitäten ist `maxParticipants` 0;
 * wer je Gruppe nachrückt, muss die Anzahl selbst ausrechnen und
 * `onlyWithPreferredType` setzen (siehe CLAUDE.md).
 * Herausgelöst aus EventService; dort steht ein Delegations-Stub.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';
import { REG_LIST_NAME } from '../EventService';

/**
 * Ersten Warteliste-Teilnehmer nachrücken: Status -> Angemeldet.
 * Wenn inheritStarterType übergeben wird (B2Run Split-Capacity), wird dieser Typ
 * dem Nachrücker zugewiesen (er erbt den Platz des Abgemeldeten).
 *
 * Wird **client-seitig** ausgeführt (von der App beim Abmelden), damit der
 * Power Automate DEX_IDReorder-Flow keinen doppelten Nachrück-Versuch macht.
 * Liefert den nachgerückten Teilnehmer (Email + Name) zurück für die
 * Nachrück-E-Mail.
 *
 * Schutz gegen Überbuchung: Wenn maxParticipants gesetzt ist und die Anzahl
 * der aktuell Angemeldeten (nach der Abmeldung) >= maxParticipants ist, wird
 * NICHT nachgerückt. Das verhindert, dass nach einer früheren Überbuchung
 * der Abbruch der Abmeldung nicht zu einer weiteren Überbuchung führt.
 */
export async function promoteFirstWaitlistItem(
  svc: EventService,
  subsiteUrl: string,
  inheritStarterType?: string,
  maxParticipants?: number,
  /** Seit v6.5: Bei B2Run-Events mit getrennten Durchstarter-/Funstarter-Wartelisten
   * hier den freigewordenen Starter-Typ mitgeben — dann wird NUR der erste
   * Warteliste-Teilnehmer mit passendem PreferredStarterType nachgerückt.
   * Wenn leer: Default-Verhalten (beliebiger Warteliste-Teilnehmer). */
  onlyWithPreferredType?: string,
  /** v17.15: Audit-Tracking — wenn der Promote durch das Cancel einer
   *  konkreten Person ausgelöst wurde (in der App-Pfad), die E-Mail
   *  und Item-Id dieser Person mitgeben. Wird auf der nachrückenden
   *  Person als ReplacedParticipantEmail + PromotedDate gespeichert,
   *  und zusätzlich auf der cancelnden Person als
   *  ReplacedByParticipantEmail (zweite MERGE-PATCH). */
  replacedByCancel?: { itemId: number; participantEmail: string },
): Promise<{ success: boolean; email?: string; name?: string; itemId?: number; skippedOverbooked?: boolean }> {
  try {
    // Überbuchungs-Schutz: Nur nachrücken, wenn tatsächlich ein Platz frei ist.
    // Bei unlimited (maxParticipants === 0 oder undefined) immer nachrücken.
    //
    // WICHTIG: '>' statt '>='. Die Abmeldung (Status->Abgemeldet) ist kurz vor
    // diesem Call passiert; falls SharePoint den Statuswechsel noch nicht in
    // getRegistrationCount reflektiert (stale read), würden wir bei einem
    // vollen Event (z.B. 128/128) mit '>=' fälschlich skippen. Mit '>' ist
    // 'registered == max' noch erlaubt (= genau ein Platz wird nachgerückt),
    // und eine echte Überbuchung (401 > 128) wird weiterhin abgefangen.
    if (maxParticipants && maxParticipants > 0) {
      const counts = await svc.getRegistrationCount(subsiteUrl);
      if (counts.registered > maxParticipants) {
        console.warn(`[DEX] promoteFirstWaitlistItem: skipping promotion - event is overbooked (${counts.registered}/${maxParticipants} registered).`);
        return { success: false, skippedOverbooked: true };
      }
    }

    // v12.10: Nachrück-Sortierung jetzt nach TeilnehmerID asc statt
    // RegistrationDate. Hintergrund: nach dem IDReorder-Flow sind die
    // TeilnehmerIDs durchlaufend (1..N aktiv, N+1.. Warteliste). Wenn
    // also Platz 100 frei wird, soll TID 101 (= erster auf der Liste)
    // nachrücken — unabhängig davon, ob TID 103 zeitlich gesehen vor
    // TID 101 registriert war (z.B. nach Re-Registration oder Wechsel
    // der Gruppe). RegistrationDate sortierte chronologisch, was bei
    // umverteilten IDs zur falschen Reihenfolge führte.
    // Bei B2Run-Split-Kapazitäten: nur die passende Warteliste durchsuchen
    // (PreferredStarterType == onlyWithPreferredType).
    let filter = `Status eq 'Warteliste'`;
    if (onlyWithPreferredType) {
      const esc = onlyWithPreferredType.replace(/'/g, "''");
      filter += ` and PreferredStarterType eq '${esc}'`;
    }
    const resp = await svc._sp.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=${encodeURIComponent(filter)}&$orderby=TeilnehmerID asc&$top=1`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return { success: false };
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    if (items.length === 0) return { success: false };

    const firstWaiting = items[0];
    // v18.71: Kern-Update (Status -> Angemeldet, ggf. StarterType) STRIKT
    // getrennt von den optionalen Audit-Feldern. Hintergrund: bei Legacy-
    // Teilnehmerlisten, die noch nie per „Spalten fixen" aktualisiert wurden,
    // fehlt die Spalte PromotedDate (erst seit v17.15). Wenn PromotedDate im
    // selben MERGE-Body steht, lehnt SharePoint den GESAMTEN Request mit
    // HTTP 400 ab („The property 'PromotedDate' does not exist…") — der
    // Nachrück-Status wird dann gar nicht gesetzt und der Button „macht
    // nichts". Deshalb zuerst nur die Pflichtfelder schreiben.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mergeBody: Record<string, any> = { 'Status': 'Angemeldet' };
    // v19.14 AUDIT-FIX: StarterType IMMER mitsetzen — entweder den vom
    // Abgemeldeten geerbten Typ (inheritStarterType, Admin-Cancel-Pfad) ODER,
    // falls keiner mitgegeben wurde (z.B. der manuelle „Nachrücken"-Button, der
    // inheritStarterType=undefined übergibt), den EIGENEN Wunsch der
    // nachrückenden Person (PreferredStarterType). Vorher blieb StarterType bei
    // Promotes ohne inheritStarterType auf Split-Events LEER → es entstanden
    // angemeldete „typlose" Personen (Audit-Befund: Andreas Jehle), die aus den
    // Gruppen-Zahlen fielen und als „Wunsch: …" angezeigt wurden. Bei
    // Nicht-Split-Events ist PreferredStarterType leer → StarterType bleibt leer
    // (korrekt, da es dort keine Gruppen gibt).
    const effectiveStarter = inheritStarterType || firstWaiting.PreferredStarterType || '';
    if (effectiveStarter) {
      mergeBody['StarterType'] = effectiveStarter;
    }
    const mergeResp = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${firstWaiting.Id})`,
      mergeBody
    );
    if (!(mergeResp.ok || mergeResp.status === 406)) return { success: false };

    // v17.15: Nachrück-Audit auf der promoteten Person — best-effort, in
    // einem SEPARATEN MERGE, damit eine fehlende Audit-Spalte (Legacy-Liste)
    // den eigentlichen Promote oben nicht kaputtmacht.
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const auditBody: Record<string, any> = { 'PromotedDate': new Date().toISOString() };
      if (replacedByCancel && replacedByCancel.participantEmail) {
        auditBody['ReplacedParticipantEmail'] = replacedByCancel.participantEmail;
      }
      await svc._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${firstWaiting.Id})`,
        auditBody
      );
    } catch (err) {
      console.warn('[DEX] promoteFirstWaitlistItem: Nachrück-Audit (PromotedDate) konnte nicht geschrieben werden — Spalte fehlt evtl. auf einer Legacy-Liste:', err);
    }

    const vorname = firstWaiting.Vorname || '';
    const nachname = firstWaiting.Nachname || '';
    const name = (vorname && nachname) ? `${vorname} ${nachname}` : (firstWaiting.ParticipantName || '');
    const email = firstWaiting.ParticipantEmail || firstWaiting.Title || '';

    // v17.15: zweite PATCH auf die cancelnde Person — „Ersetzt durch".
    if (replacedByCancel && replacedByCancel.itemId && email) {
      try {
        await svc._merge(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${replacedByCancel.itemId})`,
          { 'ReplacedByParticipantEmail': email }
        );
      } catch (err) {
        console.warn('[DEX] Nachrück-Audit auf cancelnder Person fehlgeschlagen:', err);
      }
    }

    console.warn(`[DEX] promoteFirstWaitlistItem: promoted ${name} <${email}> (item ${firstWaiting.Id}) to Angemeldet.`);
    return { success: true, email, name, itemId: firstWaiting.Id };
  } catch {
    return { success: false };
  }
}
