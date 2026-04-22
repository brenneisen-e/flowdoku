import * as React from 'react';
import { ManualSection } from '../types';
import { Callout, ClickPath } from '../ManualMockups';

export function waitlistSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'waitlist',
    title: isDe ? 'Wartelisten & Nachrücken' : 'Waitlists & Promotion',
    category: 'organizer',
    description: isDe
      ? 'Wie Wartelisten funktionieren, wie das automatische Nachrücken abläuft und was Organizer/Admins beim Abmelden sehen. Inkl. getrennte Wartelisten pro Starter-Typ bei Lauf-Events (B2Run).'
      : 'How waitlists work, how automatic promotion happens, and what organizers/admins see when cancelling. Includes split waitlists per starter type for running events (B2Run).',
    visibleFor: ['User', 'Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'user',
        steps: [
          {
            number: 1,
            title: isDe ? 'Anmeldung bei vollem Event' : 'Registering for a full event',
            description: (
              <>
                {isDe
                  ? 'Wenn du dich für ein Event anmeldest, das bereits voll ist, landest du automatisch auf der Warteliste. Du bekommst eine Bestätigungs-Mail mit dem Hinweis, dass du auf der Warteliste stehst. Sobald sich jemand abmeldet, rückst du automatisch nach und bekommst eine Nachrück-Mail plus eine Outlook-Einladung. Das Nachrücken läuft über unseren Power-Automate-Flow und erfolgt innerhalb von ~1 Minute nach der Abmeldung.'
                  : 'When you register for an event that is already full, you are automatically put on the waitlist. You receive a confirmation email mentioning the waitlist status. As soon as someone cancels, you are automatically promoted and receive a promotion email plus an Outlook invitation. Promotion runs via our Power Automate flow and typically happens within ~1 minute after the cancellation.'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'B2Run: Durchstarter oder Funstarter wählen' : 'B2Run: Choose Durchstarter or Funstarter',
            description: (
              <>
                {isDe
                  ? 'Bei Lauf-Events mit getrennten Kapazitäten (z.B. B2Run) wählst du bei der Anmeldung explizit deinen Wunsch-Starter-Typ. Wenn dein Wunsch-Typ voll ist, aber der andere Typ noch freie Plätze hat, bekommst du einen Dialog angeboten: entweder als Alternative starten oder auf die eigene Warteliste für den Wunsch-Typ. Du entscheidest selbst — kein stiller Auto-Wechsel mehr.'
                  : 'For running events with split capacities (e.g. B2Run) you pick your preferred starter type during registration. If your preferred type is full but the other has capacity, a dialog offers you two choices: start as the alternative type, or join the waitlist for your preferred type. You decide — no silent auto-switch.'}
              </>
            ),
            mockup: <Callout variant="info">{isDe ? 'Beispiel: 10 Durchstarter-Plätze voll, 5 Funstarter-Plätze frei → Dialog fragt "Als Funstarter starten oder auf Durchstarter-Warteliste?"' : 'Example: 10 Durchstarter slots full, 5 Funstarter slots open → dialog asks "Start as Funstarter or join Durchstarter waitlist?"'}</Callout>,
          },
          {
            number: 3,
            title: isDe ? 'Getrennte Wartelisten pro Typ' : 'Separate waitlists per type',
            description: (
              <>
                {isDe
                  ? 'Bei Split-Kapazitäten gibt es zwei unabhängige Wartelisten: eine für Durchstarter und eine für Funstarter. Wenn sich ein Durchstarter abmeldet, rückt nur der nächste Durchstarter-Warteliste-Teilnehmer nach — und nicht etwa ein Funstarter, der sich gerne umstellen würde. Dein gewünschter Starter-Typ wird dabei respektiert.'
                  : 'With split capacities there are two independent waitlists — one per type. When a Durchstarter cancels, only the next Durchstarter waitlist entry is promoted (not a Funstarter who would happily switch). Your preferred starter type is always respected.'}
              </>
            ),
          },
        ],
      },
      {
        perspective: 'organizer',
        steps: [
          {
            number: 1,
            title: isDe ? 'Admin-Abmeldung im Admin-Center' : 'Admin cancellation in the admin center',
            description: (
              <>
                {isDe
                  ? 'Wenn du im Admin-Center einen Teilnehmer abmeldest (roter "Abmelden"-Button in der Teilnehmer-Liste), passiert Folgendes:'
                  : 'When you cancel a participant in the admin center (red "Cancel" button in the participant list), the following happens:'}
                <ol style={{ marginTop: 8, paddingLeft: 20 }}>
                  <li>{isDe ? 'Orange Toast oben rechts erscheint: "Abmeldung von X wird verarbeitet…" mit Spinner.' : 'Orange toast top-right: "Cancellation of X is being processed…" with spinner.'}</li>
                  <li>{isDe ? 'Teilnehmer wird in der Liste auf "Abgemeldet" gesetzt + Abmelde-Mail in Queue + Outlook-Ausladen in Queue.' : 'Participant is set to "Cancelled" + cancellation email queued + Outlook removal queued.'}</li>
                  <li>{isDe ? 'Client-seitiger Promote: der erste passende Warteliste-Teilnehmer (bei B2Run-Split: gleicher Typ wie der Abgemeldete) wird nachgerückt + Nachrück-Mail + Outlook-Einladung geschickt.' : 'Client-side promote: the first matching waitlist entry (same starter type for B2Run-split) is promoted + promotion email + Outlook invite queued.'}</li>
                  <li>{isDe ? 'Grüner Toast ersetzt den orangen: "Nachgerückt: <Name> <Email> (<Starter-Typ>)". Bleibt stehen bis du auf × klickst.' : 'Green toast replaces the orange one: "Promoted: <name> <email> (<type>)". Stays until you click ×.'}</li>
                  <li>{isDe ? 'Wenn niemand auf der Warteliste steht (oder kein passender Typ): grauer Toast "Keiner zum Nachrücken — der Platz bleibt frei".' : 'If nobody is on the waitlist (or no matching type): grey toast "No one to promote — slot stays open".'}</li>
                  <li>{isDe ? 'IDReorder-Flow wird wie gewohnt getriggert → sortiert die TeilnehmerIDs im Hintergrund (1..N Angemeldete, N+1..N+M Warteliste).' : 'IDReorder flow is triggered as usual → sorts participant IDs in the background (1..N registered, N+1..N+M waitlist).'}</li>
                </ol>
              </>
            ),
            mockup: <ClickPath label="Admin" label2={isDe ? 'Teilnehmer abmelden' : 'Cancel participant'} />,
          },
          {
            number: 2,
            title: isDe ? 'User-Self-Cancel: Flow übernimmt komplett' : 'User self-cancel: flow takes over fully',
            description: (
              <>
                {isDe
                  ? 'Wenn ein User sich selbst unter "My Events" abmeldet, erfolgt KEIN client-seitiger Promote. Stattdessen übernimmt der DEX_IDReorder_TeilnehmerIDs-Flow das Nachrücken innerhalb von ~1 Minute. Das ist Absicht: der User sieht durch die Item-Level-Security auf der Teilnehmerliste ohnehin nur seinen eigenen Eintrag, daher ist die kurze Verzögerung nicht sichtbar. Nur der Admin braucht sofortiges Feedback.'
                  : 'When a user cancels themselves under "My Events", NO client-side promote happens. Instead, the DEX_IDReorder_TeilnehmerIDs flow handles the promotion within ~1 minute. This is intentional: item-level security on the participant list means the user only sees their own entry anyway, so the short delay is invisible to them. Only the admin needs immediate feedback.'}
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Split-Wartelisten im Admin-Center' : 'Split waitlists in the admin center',
            description: (
              <>
                {isDe
                  ? 'Bei B2Run-Events mit aktivierter Split-Kapazität zeigt die Teilnehmer-Tabelle eine zusätzliche Spalte "Startblock" mit dem tatsächlichen StarterType. Wenn der Teilnehmer via Fallback-Dialog einen anderen Typ genommen hat, wird sein Wunsch in Klammern daneben angezeigt (z.B. "Funstarter (Wunsch: Durchstarter)"). Die Warteliste-Sektion zeigt drei getrennte Tabellen: Durchstarter-Warteliste, Funstarter-Warteliste, "ohne Typ" (Legacy-Fallback).'
                  : 'For B2Run events with split capacity enabled, the participant table shows an extra "Startblock" column with the actual StarterType. If the participant took a different type via the fallback dialog, their preferred type is shown in parentheses (e.g. "Funstarter (Pref: Durchstarter)"). The waitlist section displays three separate tables: Durchstarter-Waitlist, Funstarter-Waitlist, "No-type" (legacy fallback).'}
              </>
            ),
            mockup: <Callout variant="info">{isDe ? 'Die Split-Kapazitäts-Übersicht oberhalb der Teilnehmer-Liste zeigt dir sofort "Durchstarter 8/10 · Warteliste: 3" und "Funstarter 18/20 · Warteliste: 5".' : 'The split capacity overview above the participant list shows at a glance "Durchstarter 8/10 · Waitlist: 3" and "Funstarter 18/20 · Waitlist: 5".'}</Callout>,
          },
        ],
      },
      {
        perspective: 'admin',
        steps: [
          {
            number: 1,
            title: isDe ? 'Was passiert wenn ich direkt in SharePoint lösche?' : 'What happens if I delete directly in SharePoint?',
            description: (
              <>
                {isDe
                  ? 'Wenn du einen Teilnehmer direkt in der SharePoint-Teilnehmerliste löschst (ohne die App), wird KEIN DEX_IDReorder-Item geschrieben — d.h. der Flow läuft nicht, niemand rückt nach, und die TeilnehmerIDs bleiben lückenhaft. Lösung: im Admin-Center einmal auf "IDs neu vergeben" klicken (macht client-seitig einen Zwei-Pass-Reorder mit Angemeldeten zuerst, Warteliste danach).'
                  : 'If you delete a participant directly in the SharePoint participant list (bypassing the app), NO DEX_IDReorder item is written — the flow does not run, nobody is promoted, and participant IDs are left with gaps. Fix: click "Reorder IDs" in the admin center (does a client-side two-pass reorder: registered first, waitlist after).'}
              </>
            ),
            warning: isDe
              ? 'Direkte SP-Änderungen vermeiden wenn möglich — die App ist die einzige Source-of-Truth für den Flow-Trigger.'
              : 'Avoid direct SP edits if possible — the app is the single source of truth for flow triggering.',
          },
          {
            number: 2,
            title: isDe ? 'Flow-Failure erkennen und beheben' : 'Detect and fix flow failures',
            description: (
              <>
                {isDe
                  ? 'Wenn der DEX_IDReorder-Flow bei einem Eintrag fehlschlägt (z.B. Graph-API-Timeout), wird der DEX_IDReorder-Listen-Eintrag auf Status="Failed" gesetzt. In dem Fall: (a) prüfe den Flow-Run-Verlauf in Power Automate für Details, (b) klicke im Admin-Center auf "IDs neu vergeben" um den Zwei-Pass-Reorder manuell zu triggern, (c) falls Nachrücken fehlt: einen Warteliste-Teilnehmer manuell per "Anmelden"-Button aktivieren.'
                  : 'If the DEX_IDReorder flow fails on an entry (e.g. Graph API timeout), the DEX_IDReorder list item is set to Status="Failed". Action: (a) check the flow run history in Power Automate for details, (b) click "Reorder IDs" in the admin center to trigger the two-pass reorder manually, (c) if promotion is missing: manually activate a waitlist entry via "Register"-button.'}
              </>
            ),
          },
        ],
      },
    ],
  };
}
