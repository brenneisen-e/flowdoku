/**
 * v30.66 — Modularisierung Stufe 2: Thema „Fun-Zone/Quiz" — der Punktestand
 * einer Person wird auf ihrer Anmeldezeile gespeichert (QuizScore/QuizAnswers),
 * die Spalten dafür entstehen erst beim ersten Speichern.
 * Herausgelöst aus EventService; dort steht ein Delegations-Stub.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';
import { REG_LIST_NAME } from '../EventService';

/**
 * Quiz-Fortschritt in die Registrierung eines Teilnehmers schreiben.
 *
 * - answers: ausgewählte Antwort-Indices pro Frage (Array von Arrays, weil
 *   Fragen mehrere richtige Antworten haben können). Unbeantwortete Fragen
 *   bleiben als leeres Array `[]` stehen, damit der Index-Offset erhalten bleibt.
 * - score: aktuell erreichte Punkte (Anzahl korrekt beantworteter Fragen).
 * - isComplete: true wenn alle Fragen beantwortet sind — dann wird auch
 *   `QuizCompletedAt` gesetzt. Andernfalls bleibt QuizCompletedAt unverändert
 *   (null/leer), sodass der Teilnehmer als "teilweise beantwortet" gelistet wird.
 *
 * Ersetzt die früher nur-am-Ende aufgerufene `saveQuizResult()`. Wird jetzt
 * bei jedem "Weiter"-Klick im QuizPlayer aufgerufen (Auto-Save), damit der
 * Teilnehmer beim späteren Wiedereintritt an derselben Stelle weitermachen kann.
 */
export async function saveQuizProgress(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  score: number,
  answers: number[][],
  isComplete: boolean
): Promise<boolean> {
  try {
    // Vor dem Schreiben sicherstellen, dass die Quiz-Spalten auf der
    // Teilnehmer-Liste existieren. Bei Bestandsevents (vor Quiz-Feature
    // angelegt) fehlen sie oft; _merge mit odata=nometadata schluckt
    // unbekannte Felder stumm und das Save wirkt wie "gespeichert",
    // persistiert aber nichts.
    // Silent: wenn der aktuelle User keine Manage-Lists-Permission auf
    // der Subsite hat, schlägt das Anlegen fehl (Regular User). Dann
    // kann nur ein Admin/Organizer die Spalten anlegen — dafür gibt's
    // die "Spalten fixen"-Funktion im Admin Center.
    await ensureQuizColumnsOnRegList(svc, subsiteUrl);

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const payload: { [key: string]: any } = {
      'QuizScore': score,
      'QuizAnswers': JSON.stringify(answers),
    };
    if (isComplete) {
      payload.QuizCompletedAt = new Date().toISOString();
    }
    const resp = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      payload
    );
    return resp.ok || resp.status === 406;
  } catch {
    return false;
  }
}

/**
 * Quiz-Spalten auf der Teilnehmer-Liste einer Event-Subsite anlegen,
 * falls sie fehlen. Idempotent und silent: bei fehlender Permission
 * kein Crash, einfach kein-op.
 */
async function ensureQuizColumnsOnRegList(svc: EventService, subsiteUrl: string): Promise<void> {
  try {
    const fieldsResp = await svc._sp.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields?$filter=Hidden eq false&$top=200&$select=InternalName`,
      SPHttpClient.configurations.v1
    );
    if (!fieldsResp.ok) return;
    const fieldsData = await fieldsResp.json();
    const existing = new Set<string>(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (fieldsData.value || fieldsData.d?.results || []).map((f: any) => f.InternalName)
    );
    const required: Array<{ title: string; type: number }> = [
      { title: 'QuizScore', type: 9 },      // Number
      { title: 'QuizAnswers', type: 3 },    // Note (multiline)
      { title: 'QuizCompletedAt', type: 4 } // DateTime
    ];
    for (const f of required) {
      if (existing.has(f.title)) continue;
      try {
        await svc._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`,
          {
            '__metadata': { 'type': 'SP.Field' },
            'Title': f.title,
            'FieldTypeKind': f.type,
            'Required': false,
          }
        );
        console.warn(`[DEX] ensureQuizColumnsOnRegList: ${f.title} nachgelegt auf ${subsiteUrl}`);
      } catch {
        // keine Permission -> silent. User braucht Admin der "Spalten fixen" macht.
      }
    }
  } catch { /* silent */ }
}
