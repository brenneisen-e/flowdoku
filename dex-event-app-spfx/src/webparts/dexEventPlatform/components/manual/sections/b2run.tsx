import * as React from 'react';
import { ManualSection } from '../types';
import { Callout } from '../ManualMockups';
import { AppPreview } from '../previews/AppPreview';
import { DEMO_EVENT_ID } from '../previews/PreviewProviders';
import Header from '../../Header';
import EventCreationPage from '../../EventCreationPage';

export function b2runSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'b2run',
    title: isDe ? 'Lauf-Events (B2Run) einrichten' : 'Running events (B2Run) setup',
    category: 'organizer',
    description: isDe
      ? 'Komplett-Guide für B2Run-ähnliche Lauf-Events mit getrennten Starter-Typen (Durchstarter / Funstarter), eigenen Kapazitäten, getrennten Wartelisten, typ-bewusstem Nachrücken und Trainings-Sessions als Sub-Events.'
      : 'Full guide for B2Run-style running events with split starter types (Durchstarter / Funstarter), separate capacities, separate waitlists, type-aware promotion and training sessions as sub-events.',
    keywords: 'b2run b2 run lauf laufevent lauf-event running run marathon firmenlauf staffel durchstarter funstarter starter-typ startertyp startblock startblöcke split-kapazität geteilte kapazität split capacity getrennte warteliste waitlist nachrücken promotion leistungsnachweis proof of performance agb datenschutz zimmerpartner roommate laufshirt trainings-session training sub-event subevent',
    visibleFor: ['Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'organizer',
        steps: [
          {
            number: 1,
            title: isDe ? 'Warum Split-Kapazitäten?' : 'Why split capacities?',
            description: (
              <>
                {isDe
                  ? 'Bei Lauf-Events gibt es typischerweise zwei unterschiedlich schnelle Starter-Gruppen — die "Durchstarter" (ambitioniert, streben Bestzeit an) und die "Funstarter" (entspannter, Hauptsache Mitmachen). Organisatoren vergeben oft getrennte Kontingente, z.B. 10 Durchstarter + 90 Funstarter — damit der Spaß-Startblock nicht versehentlich von Ambitionierten gefüllt wird oder umgekehrt. Die DEX-App unterstützt das mit einer expliziten Split-Kapazitäts-Funktion: pro Event können zwei getrennte Kapazitäten und zwei getrennte Wartelisten verwaltet werden.'
                  : 'Running events typically have two differently-paced starter groups — "Durchstarter" (ambitious, aiming for best times) and "Funstarter" (relaxed, focused on participation). Organizers usually assign separate quotas, e.g. 10 Durchstarter + 90 Funstarter — so the fun block is not accidentally filled with ambitious runners or vice versa. The DEX app supports this with an explicit split-capacity feature: each event can have two separate capacities and two separate waitlists.'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Event anlegen — Split-Kapazität aktivieren' : 'Create event — enable split capacity',
            description: (
              <>
                {isDe
                  ? 'Beim Event-Anlegen im Schritt "Kapazität & Sichtbarkeit" (Schritt 5) findest du die Checkbox "Lauf-Event mit getrennten Starter-Kapazitäten". Aktivieren — und die klassische "Max. Teilnehmer"-Eingabe wird ersetzt durch zwei Felder: Durchstarter-Kapazität und Funstarter-Kapazität. Die Summe ist automatisch die Gesamt-Kapazität.'
                  : 'When creating an event, in the "Capacity & visibility" step (step 5) you find the checkbox "Running event with split starter capacities". Activate it — and the classic "Max participants" input is replaced by two fields: Durchstarter capacity and Funstarter capacity. The sum is automatically the total capacity.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Wizard Schritt 5: Split-Kapazitäten (echte Ansicht)' : 'Wizard step 5: Split capacities (real view)'}
                role="Organizer"
                page="edit-event"
                selectedEventId={DEMO_EVENT_ID}
                width={1024}
                device="laptop"
                initialStep={4}
              >
                <Header />
                <EventCreationPage />
              </AppPreview>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Teilnehmer wählt den Wunsch-Starter-Typ' : 'Participant picks preferred starter type',
            description: (
              <>
                {isDe
                  ? 'Auf der Registrierungs-Seite erscheint bei Split-Events automatisch eine Auswahl "Durchstarter vs. Funstarter". Der User wählt seinen Wunsch-Typ. Die App prüft dann:'
                  : 'On the registration page, split events automatically show a "Durchstarter vs. Funstarter" selection. The user picks their preferred type. The app then checks:'}
                <ol style={{ marginTop: 8, paddingLeft: 20 }}>
                  <li>{isDe ? 'Wunsch-Typ hat freie Plätze → direkt angemeldet (Status=Angemeldet, StarterType=Wunsch).' : 'Preferred type has free slots → registered directly (Status=Angemeldet, StarterType=Preferred).'}</li>
                  <li>{isDe ? 'Wunsch-Typ voll, Alt-Typ frei → Modal-Dialog fragt den User: "Als Alt-Typ starten oder auf Wunsch-Warteliste?" — keine stille Umschaltung mehr.' : 'Preferred type full, alt type open → modal dialog asks the user: "Start as alt or join preferred waitlist?" — no silent auto-switch.'}</li>
                  <li>{isDe ? 'Beide voll → Warteliste für den Wunsch-Typ (PreferredStarterType wird gespeichert, damit der Nachrück-Flow später den richtigen Typ trifft).' : 'Both full → waitlist for the preferred type (PreferredStarterType is saved so the promotion flow later matches the correct type).'}</li>
                </ol>
              </>
            ),
          },
          {
            number: 4,
            title: isDe ? 'Getrennte Wartelisten & typ-bewusstes Nachrücken' : 'Separate waitlists & type-aware promotion',
            description: (
              <>
                {isDe
                  ? 'Wenn ein Durchstarter sich abmeldet, rückt nur ein Durchstarter-Warteliste-Teilnehmer nach — NICHT ein Funstarter, der sich eventuell umstellen würde. Der PreferredStarterType jedes Warteliste-Teilnehmers wird respektiert. Dieser Mechanismus läuft sowohl im Client (bei Admin-Cancels, sofort) als auch im Power-Automate-Flow (bei User-Self-Cancels, innerhalb ~1 Minute). In der Admin-UI findest du drei separate Warteliste-Tabellen (Durchstarter, Funstarter, "ohne Typ" als Legacy-Fallback).'
                  : 'When a Durchstarter cancels, only a Durchstarter waitlist entry is promoted — NOT a Funstarter who might be willing to switch. The PreferredStarterType of each waitlist entry is respected. This mechanism runs both in the client (on admin cancels, immediately) and in the Power Automate flow (on user self-cancels, within ~1 minute). The admin UI shows three separate waitlist tables (Durchstarter, Funstarter, "No type" legacy fallback).'}
              </>
            ),
            mockup: <Callout variant="info">{isDe ? 'Beispiel: 10 Durch (voll) / 90 Fun (voll). Durchstarter-Warteliste: 5, Funstarter-Warteliste: 20. Ein Durchstarter meldet sich ab → der oberste Durchstarter-Warteliste-Teilnehmer rückt nach, StarterType wird automatisch auf Durchstarter gesetzt. Die Funstarter-Warteliste bleibt unberührt.' : 'Example: 10 Durch (full) / 90 Fun (full). Durchstarter waitlist: 5, Funstarter waitlist: 20. A Durchstarter cancels → the top Durchstarter waitlist entry is promoted, StarterType is auto-set to Durchstarter. The Funstarter waitlist is untouched.'}</Callout>,
          },
          {
            number: 5,
            title: isDe ? 'Admin-Center: Split-Übersicht + Startblock-Spalte' : 'Admin center: split overview + starter-block column',
            description: (
              <>
                {isDe
                  ? 'Im Admin-Center erscheint oberhalb der Teilnehmer-Liste eine Übersicht mit zwei Kacheln — eine pro Starter-Typ mit aktueller Belegung und Warteliste-Zahl (z.B. "Durchstarter 8/10 · Warteliste: 3"). In der Teilnehmer-Tabelle zeigt eine zusätzliche Spalte "Startblock" den tatsächlichen StarterType. Wenn ein Teilnehmer über den Fallback-Dialog einen anderen Typ genommen hat als ursprünglich gewünscht, wird der Wunsch in Klammern angezeigt (z.B. "Funstarter (Wunsch: Durchstarter)").'
                  : 'In the admin center, an overview appears above the participant list with two cards — one per starter type showing current occupancy and waitlist count (e.g. "Durchstarter 8/10 · Waitlist: 3"). In the participant table an additional "Startblock" column shows the actual StarterType. If a participant chose a different type via the fallback dialog than originally preferred, the preferred type is shown in parentheses (e.g. "Funstarter (Pref: Durchstarter)").'}
              </>
            ),
          },
          {
            number: 6,
            title: isDe ? 'Trainings-Sessions als Sub-Events' : 'Training sessions as sub-events',
            description: (
              <>
                {isDe
                  ? 'Bei B2Run-Events bieten Organizer oft optionale Trainings-Sessions in den Wochen vor dem Lauf an. Die DEX-App löst das über Sub-Events: im Schritt "Sub-Events" (Schritt 3) des Event-Editors kannst du pro Session einen Eintrag mit Titel, Ort, Start/Ende, Anmeldeschluss und Kapazität anlegen. Sub-Events werden als eigenständige DEX_Events-Items mit parentEventId verwaltet — sie haben eine eigene Teilnehmerliste, einen eigenen Outlook-Termin und eigene Mails. Der Teilnehmer meldet sich zuerst für das Haupt-Event an, dann auf dem Success-Screen für beliebige Sessions.'
                  : 'B2Run organizers often offer optional training sessions in the weeks before the run. The DEX app solves this via sub-events: in the "Sub-events" step (step 3) of the event editor you can add one entry per session with title, location, start/end, registration cutoff and capacity. Sub-events are managed as standalone DEX_Events items with parentEventId — they have their own participant list, their own Outlook appointment and their own emails. The participant first registers for the main event, then on the success screen for any sessions.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Wizard Schritt 3: Sub-Events (echte Ansicht)' : 'Wizard step 3: Sub-events (real view)'}
                role="Organizer"
                page="edit-event"
                selectedEventId={DEMO_EVENT_ID}
                width={1024}
                device="laptop"
                initialStep={2}
              >
                <Header />
                <EventCreationPage />
              </AppPreview>
            ),
          },
          {
            number: 7,
            title: isDe ? 'Startblock-Custom-Field (optional)' : 'Starter block custom field (optional)',
            description: (
              <>
                {isDe
                  ? 'Wenn das B2Run-Template gewählt wird, legt die App automatisch Custom-Fields für Startblock-Auswahl, AGB-/Datenschutz-Zustimmung, Zimmerpartner usw. an. Diese Felder können im Schritt "Felder" (Schritt 6) angepasst werden. Alle Antworten landen in der Teilnehmerliste als SP-Spalten.'
                  : 'When the B2Run template is selected, the app automatically creates custom fields for starter-block selection, terms & privacy acceptance, roommate picker etc. These fields can be adjusted in the "Fields" step (step 6). All answers are stored in the participant list as SP columns.'}
              </>
            ),
          },
          {
            number: 8,
            title: isDe ? 'Starter-Typ ↔ Startblock verknüpfen' : 'Link starter type ↔ start block',
            description: (
              <>
                {isDe
                  ? 'Bei aktivierter Split-Kapazität kann der Organizer direkt im Schritt "Kapazität & Sichtbarkeit" (Schritt 5) jedem Starter-Typ einen Startblock zuordnen — z.B. "Durchstarter → Block A", "Funstarter → Block B". Wenn gesetzt, wird das Startblock-Custom-Field bei der Registrierung automatisch anhand des gewählten Starter-Typs gesetzt. Der User muss den Block dann nicht extra auswählen; das b2run_startblock-Feld wird in der Registrierungs-Maske ausgeblendet. Die freien Plätze pro Starter-Typ entsprechen 1:1 den freien Plätzen im zugeordneten Block.'
                  : 'With split capacity enabled the organizer can map a start block to each starter type directly in the "Capacity & visibility" step (step 5) — e.g. "Durchstarter → Block A", "Funstarter → Block B". If set, the start block custom field is auto-populated at registration based on the chosen starter type. The user does not have to pick the block separately; the b2run_startblock field is hidden in the registration form. Free seats per starter type map 1:1 to free seats in the associated block.'}
              </>
            ),
          },
          {
            number: 9,
            title: isDe ? 'Leistungsnachweis-Pflicht für Durchstarter' : 'Proof of performance for Durchstarter',
            description: (
              <>
                {isDe
                  ? 'Ebenfalls im Schritt "Kapazität & Sichtbarkeit" (Schritt 5) findest du die Option "Leistungsnachweis für Durchstarter erforderlich". Wenn aktiviert, muss der User beim Wählen von Durchstarter eine zusätzliche Pflicht-Checkbox bestätigen ("Ich bestätige, dass ein entsprechender Leistungsnachweis vorliegt"). Ohne Bestätigung wird die Anmeldung blockiert. Die Bestätigung landet als b2run_leistungsnachweis-Eintrag in den CustomData der Teilnehmerregistrierung, sodass Organizer später kontrollieren können, wer bestätigt hat.'
                  : 'Also in the "Capacity & visibility" step (step 5) you find the option "Proof of performance required for Durchstarter". When enabled, the user must tick an additional mandatory checkbox when selecting Durchstarter ("I confirm that a valid proof of performance exists"). Without confirmation, the registration is blocked. The confirmation lands as b2run_leistungsnachweis entry in the CustomData of the participant registration so organizers can later verify who confirmed.'}
              </>
            ),
            mockup: <Callout variant="info">{isDe ? 'Typische Anwendung: Organizer vergibt 10 Durchstarter-Plätze an schnelle Läufer. Mit der Leistungsnachweis-Pflicht bestätigt jeder Durchstarter-Anmelder explizit, dass er die Qualifikation hat — dient als Gatekeeping-Mechanismus und rechtliche Absicherung.' : 'Typical use case: organizer allocates 10 Durchstarter slots to fast runners. With the proof requirement, every Durchstarter applicant explicitly confirms having the qualification — acts as gatekeeping mechanism and legal safeguard.'}</Callout>,
          },
        ],
      },
    ],
  };
}
