/**
 * v32.3: Kurzantworten auf Standardfragen — EINE Quelle für die globale Suche
 * („Kurz beantwortet", seit v32.0.3) und das Fragen-Modal „Hast du Fragen?".
 * Nutzer-Ansage 28.09.2026: „Ich brauche einen Agenten, der Standardfragen /
 * Stichwörter gut in einer kurzen Antwort ausformuliert — also ‚Reminder' →
 * ‚Ja, du kannst Reminder für dein Event verschicken. Das geht wie folgt: …'".
 *
 * Kein Sprachmodell: Die App hat im SharePoint keins. Die Antworten sind
 * kuratiert und nennen NUR Wege, die es in der App gibt — wer einen Eintrag
 * ergänzt, prüft den Klickweg vorher im Code (Aktionsname, Schritt-Titel).
 * Die Erkennung (`matchFaq`) sucht Stichwörter im freien Fragetext, auch als
 * Wortanfang („Reminder-Mail", „erinnern").
 */
import type { Page } from '../context/NavigationContext';

export type FaqGate = 'all' | 'manage' | 'admin';

export interface FaqEntry {
  key: string;
  /** Frage, wie sie über der Antwort steht. */
  de: string; en: string;
  /** Die Antwort in einem Satz — „Ja, …" / „Über …". */
  aDe: string; aEn: string;
  /** Der Klickweg als Schritte (optional). */
  stepsDe?: string[]; stepsEn?: string[];
  kw: string[];
  page: Page;
  gate: FaqGate;
}

export const FAQ_CATALOG: FaqEntry[] = [
  { key: 'reminder', de: 'Wie verschicke ich einen Reminder?', en: 'How do I send a reminder?',
    aDe: 'Ja, du kannst Reminder verschicken — an alle, die noch nicht geantwortet haben, oder an alle Angemeldeten. Beides läuft über die Aktion „E-Mail versenden“.',
    aEn: 'Yes, you can send reminders — to everyone who has not responded yet, or to everyone registered. Both use the „Send email“ action.',
    stepsDe: ['Organizer Center öffnen und dein Event wählen.', 'Oben „Aktionen“ → „E-Mail versenden“.', 'Empfänger wählen: „Erinnerung — sieht das Event, hat aber noch nicht geantwortet“ für alle, die sich noch nicht an- oder abgemeldet haben — oder „Alle aktiven Teilnehmer“ für einen Reminder an die Angemeldeten.', 'Text schreiben, erst an dich testen, dann senden.'],
    stepsEn: ['Open the Organizer Center and choose your event.', 'At the top: „Actions“ → „Send email“.', 'Pick the recipients: „Reminder — can see the event but has not responded yet“ for everyone who has neither registered nor cancelled — or „All active participants“ for a reminder to those registered.', 'Write the text, send a test to yourself, then send.'],
    kw: ['reminder', 'erinnerung', 'erinnern', 'nachfassen', 'rundmail', 'massenmail', 'nicht geantwortet', 'nicht angemeldet'], page: 'admin', gate: 'manage' },
  { key: 'invite', de: 'Wie lade ich Leute zu meinem Event ein?', en: 'How do I invite people to my event?',
    aDe: 'Über die Aktion „Einladungsmail verschicken“ — sie enthält den Anmelde-Link und geht an dich zum Weiterleiten oder direkt an einen Verteiler.',
    aEn: 'With the „Send invitation mail“ action — it contains the registration link and goes to you for forwarding or straight to a distribution list.',
    stepsDe: ['Event zuerst live schalten (sonst sieht niemand die Anmeldeseite).', 'Organizer Center → Event → „Aktionen“ → „Einladungsmail verschicken“.'],
    stepsEn: ['Publish the event first (otherwise nobody sees the registration page).', 'Organizer Center → event → „Actions“ → „Send invitation mail“.'],
    kw: ['einladung', 'einladen', 'invite', 'invitation', 'anmelde-link', 'anmeldelink', 'link'], page: 'admin', gate: 'manage' },
  { key: 'live', de: 'Wie schalte ich ein Event live?', en: 'How do I publish an event?',
    aDe: 'Im Organizer Center unter „Nächste Schritte“ → „Live schalten“. Danach sehen es die berechtigten Gruppen.',
    aEn: 'In the Organizer Center under „Next steps“ → „Go live“. The eligible groups can then see it.',
    stepsDe: ['Organizer Center → Event wählen.', '„Nächste Schritte“ → „Live schalten“ (oder im Assistenten in Schritt 1 den Entwurf-Haken entfernen).'],
    stepsEn: ['Organizer Center → choose the event.', '„Next steps“ → „Go live“ (or untick draft in step 1 of the wizard).'],
    kw: ['live', 'veröffentlichen', 'freischalten', 'aktivieren', 'entwurf', 'publish', 'sichtbar'], page: 'admin', gate: 'manage' },
  { key: 'export', de: 'Wie bekomme ich die Teilnehmerliste als Excel?', en: 'How do I get the attendee list as Excel?',
    aDe: 'Über die Aktion „Teilnehmerliste exportieren (Excel)“ im Organizer Center — mit allen Antworten aus dem Anmeldeformular.',
    aEn: 'With the „Export attendee list (Excel)“ action in the Organizer Center — including all answers from the registration form.',
    stepsDe: ['Organizer Center → Event wählen.', '„Aktionen“ → „Teilnehmerliste exportieren (Excel)“.'],
    stepsEn: ['Organizer Center → choose the event.', '„Actions“ → „Export attendee list (Excel)“.'],
    kw: ['export', 'exportieren', 'excel', 'csv', 'teilnehmerliste', 'liste herunterladen', 'download'], page: 'admin', gate: 'manage' },
  { key: 'qr', de: 'Wie funktioniert der Check-in?', en: 'How does check-in work?',
    aDe: 'Du verschickst QR-Codes über die Aktion „QR-Codes versenden“; am Event-Tag scannt das Check-in-Team unter „Check-in“. Ohne Kamera geht die Teilnehmer-ID unter jedem QR-Code.',
    aEn: 'Send QR codes with the „Send QR codes“ action; on the day the check-in team scans under „Check-in“. Without a camera, use the attendee ID below every QR code.',
    kw: ['check-in', 'checkin', 'einchecken', 'scannen', 'qr', 'qr-code', 'anwesenheit'], page: 'check-in', gate: 'manage' },
  { key: 'waitlist', de: 'Wie rücken Personen von der Warteliste nach?', en: 'How do people move up from the waitlist?',
    aDe: 'Bei jeder Abmeldung rückt automatisch die nächste Person nach. Hast du die Plätze erhöht, füllt die Aktion „Freie Plätze mit Warteliste füllen“ die neuen Plätze.',
    aEn: 'With every cancellation the next person moves up automatically. After raising the capacity, the „Fill free seats from waitlist“ action fills the new seats.',
    kw: ['warteliste', 'nachrücken', 'nachruecken', 'waitlist', 'nachrücker', 'ausgebucht'], page: 'admin', gate: 'manage' },
  { key: 'coorg', de: 'Wie füge ich weitere Organizer hinzu?', en: 'How do I add more organizers?',
    aDe: 'Im Event-Assistenten, Schritt 2 „Organizer & Team“: Personen als Organizer oder Co-Organizer eintragen und speichern. Sie bekommen eine Mail; die Rechte sind nach ein paar Minuten aktiv.',
    aEn: 'In the event wizard, step 2 „Organizers & Team“: add people as organizers or co-organizers and save. They get an email; permissions are active after a few minutes.',
    kw: ['co-organizer', 'coorganizer', 'organizer hinzufügen', 'mitorganisator', 'kollege', 'kollegin', 'rechte', 'zugriff'], page: 'admin', gate: 'manage' },
  { key: 'proxy', de: 'Kann ich jemand anderen anmelden?', en: 'Can I register someone else?',
    aDe: 'Ja — auf der Event-Übersicht über „Für andere Person registrieren“. Die Person bekommt die Bestätigung, du wirst als Anmeldende:r vermerkt.',
    aEn: 'Yes — on the event overview via „Register another person“. The person gets the confirmation; you are recorded as the one who registered them.',
    kw: ['für andere', 'andere person', 'stellvertretend', 'anmelden für', 'assistenz', 'register someone', 'kollegen anmelden'], page: 'register', gate: 'all' },
  { key: 'series', de: 'Wie lege ich eine Terminserie an?', en: 'How do I create a series?',
    aDe: 'Im Event-Assistenten, Schritt 1, unter dem Zeitraum: „Als Serie anlegen“. Jeder Termin wird ein eigenes Sub-Event mit eigener Teilnehmerliste.',
    aEn: 'In the event wizard, step 1, below the dates: „Create as series“. Every date becomes its own sub-event with its own attendee list.',
    kw: ['serie', 'terminserie', 'wiederkehrend', 'wöchentlich', 'regelmäßig', 'series', 'recurring'], page: 'create-event', gate: 'manage' },
  { key: 'position', de: 'Kann ich eine Frage nur bestimmten Positionen zeigen?', en: 'Can I show a question only to certain positions?',
    aDe: 'Ja — im Assistenten, Schritt „Fragen im Anmeldeformular“: bei der Frage „Details“ → „Nur für bestimmte Positionen anzeigen“ (z. B. alle außer Partner und Director).',
    aEn: 'Yes — in the wizard step „Registration form questions“: open the question’s „Details“ → „Show only for certain positions“ (e.g. everyone except partners and directors).',
    kw: ['position', 'positionen', 'partner', 'director', 'jobtitel', 'nur für', 'bestimmte personen'], page: 'admin', gate: 'manage' },
];

const normFaq = (s: string): string => (s || '').toLowerCase();

/**
 * Passende Kurzantworten zu einem freien Text, beste zuerst. Ein Stichwort
 * trifft, wenn es im Text steht oder ein Wort des Texts damit beginnt
 * (bzw. umgekehrt ab vier Zeichen) — „Reminder-Mail" trifft „reminder",
 * „erinnern" trifft „erinnerung" nicht, dafür steht es selbst in der Liste.
 */
export function matchFaq(text: string, canManage: boolean, max = 1): FaqEntry[] {
  const t = normFaq(text);
  if (t.trim().length < 3) return [];
  const words = t.split(/[^a-z0-9äöüß-]+/).filter(w => w.length >= 3);
  const scored: Array<{ f: FaqEntry; score: number }> = [];
  for (const f of FAQ_CATALOG) {
    if (f.gate !== 'all' && !canManage) continue;
    let score = 0;
    for (const k of f.kw) {
      const kk = normFaq(k);
      if (kk.indexOf(' ') >= 0 ? t.indexOf(kk) >= 0
        : words.some(w => w === kk || w.indexOf(kk) === 0 || (w.length >= 4 && kk.indexOf(w) === 0))) score++;
    }
    if (score > 0) scored.push({ f, score });
  }
  return scored.sort((a, b) => b.score - a.score).slice(0, max).map(x => x.f);
}
