import * as React from 'react';
import { ManualSection } from '../types';
import { Callout } from '../ManualMockups';

/**
 * v18.33 — Handbuch-Sektion „Self-Check-in per QR-Code".
 *
 * Erklärt das Aktivieren (Organizer), die beiden QR-Modi (statisches PDF +
 * rotierende Live-Anzeige) und den Scan-Vorgang aus Teilnehmer-Sicht.
 */
export function selfCheckInSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'self-check-in',
    title: isDe ? 'Self-Check-in per QR-Code' : 'Self check-in via QR code',
    category: 'general',
    description: isDe
      ? 'Teilnehmer checken sich am Veranstaltungstag selbst ein, indem sie einen event-spezifischen QR-Code mit der Handy-Kamera scannen — kein Anstehen am Check-in-Schalter, kein In-App-Scanner. Zwei Modi: druckbares QR-PDF (bequem) und rotierende Live-Anzeige (foto-sicher).'
      : 'Attendees check themselves in on event day by scanning an event-specific QR code with their phone camera — no queue at the desk, no in-app scanner. Two modes: printable QR PDF (convenient) and rotating live display (photo-safe).',
    visibleFor: ['User', 'Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'organizer',
        title: isDe ? 'Als Organizer:in' : 'As organizer',
        intro: isDe ? (
          <>
            Du aktivierst Self-Check-in <strong>pro Event</strong> im Wizard unter <strong>&bdquo;Kapazität &amp; Sichtbarkeit&ldquo;</strong>. Danach stehen im Admin Center zwei Wege bereit, den QR-Code bereitzustellen: ein <strong>druckbares PDF</strong> und eine <strong>rotierende Live-Anzeige</strong>. Beide nutzen denselben event-spezifischen Code — Teilnehmer scannen ihn mit der normalen Handy-Kamera und werden automatisch als anwesend markiert.
          </>
        ) : (
          <>
            You enable self check-in <strong>per event</strong> in the wizard under <strong>&bdquo;Capacity &amp; visibility&ldquo;</strong>. Afterwards the admin center offers two ways to provide the QR code: a <strong>printable PDF</strong> and a <strong>rotating live display</strong>. Both use the same event-specific code — attendees scan it with their normal phone camera and are automatically marked present.
          </>
        ),
        steps: [
          {
            number: 1,
            title: isDe ? 'Self-Check-in aktivieren' : 'Enable self check-in',
            description: (
              <>
                {isDe
                  ? 'Im Event-Wizard im Schritt „Kapazität & Sichtbarkeit" findest du den Schalter „Self-Check-in per QR-Code". Beim Aktivieren öffnet sich ein Erklär-Modal, das beide Modi (PDF + Live-Anzeige) beschreibt, und es wird einmalig ein geheimer Code für dieses Event erzeugt. Optional kannst du ein Zeitfenster (Von/Bis) festlegen — leer bedeutet: Check-in ist nur am Veranstaltungstag möglich.'
                  : 'In the event wizard, step „Capacity & visibility", you find the toggle „Self check-in via QR code". Enabling it opens an explainer modal describing both modes (PDF + live display) and generates a one-time secret code for this event. Optionally set a time window (from/until) — empty means check-in is only possible on the event day.'}
              </>
            ),
            tip: isDe
              ? 'Das Zeitfenster ist die einfachste Schutzmaßnahme gegen „von zu Hause einchecken" — kombiniere es immer mit dem QR-Code.'
              : 'The time window is the simplest safeguard against „checking in from home" — always combine it with the QR code.',
          },
          {
            number: 2,
            title: isDe ? 'Modus A: druckbares QR-PDF' : 'Mode A: printable QR PDF',
            description: (
              <>
                {isDe
                  ? 'Nach dem Speichern findest du im Admin Center dieses Events die Aktion „Self-Check-in: QR-PDF". Sie lädt ein A4-PDF mit dem QR-Code und einer kurzen Anleitung herunter — zum Ausdrucken und Aushängen am Eingang. Bequem, aber: ein abfotografierter Code lässt sich theoretisch weitergeben. Deshalb am besten mit dem Zeitfenster „nur am Event-Tag" kombinieren.'
                  : 'After saving, the admin center of this event offers the action „Self check-in: QR PDF". It downloads an A4 PDF with the QR code and short instructions — to print and post at the entrance. Convenient, but: a photographed code could in theory be shared. Best combined with the „event day only" window.'}
              </>
            ),
            mockup: (
              <Callout variant="info" title={isDe ? 'Wo finde ich das?' : 'Where do I find this?'}>
                {isDe
                  ? 'Admin Center → Event auswählen → Aktion „Self-Check-in: QR-PDF". Der Download startet sofort.'
                  : 'Admin center → select event → action „Self check-in: QR PDF". The download starts immediately.'}
              </Callout>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Modus B: rotierende Live-Anzeige (foto-sicher)' : 'Mode B: rotating live display (photo-safe)',
            description: (
              <>
                {isDe
                  ? 'Die Aktion „Self-Check-in: Live-Anzeige" öffnet einen großen QR-Code, der automatisch alle paar Sekunden wechselt. Zeige ihn auf einem Bildschirm am Eingang (Laptop, Tablet, Beamer) — per „Vollbild"-Button bildschirmfüllend. Ein abfotografierter Code verfällt sofort, weil er nur im aktuellen Zeitfenster gültig ist. Das ist die richtige Wahl, wenn du sicher sein willst, dass nur Anwesende einchecken.'
                  : 'The action „Self check-in: live display" opens a large QR code that changes automatically every few seconds. Show it on a screen at the entrance (laptop, tablet, projector) — fullscreen via the „Fullscreen" button. A photographed code expires instantly because it is only valid in the current time window. The right choice if you want to ensure only people on site check in.'}
              </>
            ),
            tip: isDe
              ? 'Live-Anzeige + Zeitfenster ist die sicherste Kombination. PDF ist die bequemste. Du kannst auch beides nebeneinander nutzen.'
              : 'Live display + time window is the most secure combination. PDF is the most convenient. You can use both side by side.',
          },
        ],
      },
      {
        perspective: 'user',
        title: isDe ? 'Als Teilnehmer:in' : 'As attendee',
        steps: [
          {
            number: 1,
            title: isDe ? 'QR-Code mit der Handy-Kamera scannen' : 'Scan the QR code with your phone camera',
            description: (
              <>
                {isDe
                  ? 'Am Eingang öffnest du einfach die normale Kamera-App deines Smartphones und richtest sie auf den QR-Code (auf dem Aushang oder dem Bildschirm). Es erscheint ein Link — tippe ihn an. Du brauchst KEINEN Scanner in der App: der native Kamera-Scan funktioniert zuverlässig auf jedem Handy, auch wenn die SharePoint-App keinen Kamera-Zugriff erlaubt.'
                  : 'At the entrance, simply open your phone\'s normal camera app and point it at the QR code (on the poster or the screen). A link appears — tap it. You do NOT need an in-app scanner: the native camera scan works reliably on any phone, even when the SharePoint app blocks camera access.'}
              </>
            ),
            tip: isDe
              ? 'Wichtig: mit der Handy-Kamera scannen, nicht in der App nach einem Scanner suchen. Genau das umgeht das bekannte Kamera-Problem der SharePoint-App.'
              : 'Important: scan with the phone camera, do not look for a scanner inside the app. This is exactly what works around the known camera issue of the SharePoint app.',
          },
          {
            number: 2,
            title: isDe ? 'Automatisch eingecheckt' : 'Checked in automatically',
            description: (
              <>
                {isDe
                  ? 'Der Link öffnet die DEX-App, erkennt dich automatisch über deinen Login und checkt dich ein — du siehst eine grüne Bestätigung „Eingecheckt!". Du checkst dich immer nur selbst ein; der Code lässt sich nicht für andere Personen verwenden. Falls etwas nicht passt, zeigt die App den Grund (z.B. „Keine Anmeldung gefunden", „Check-in noch nicht offen" oder „Code abgelaufen — bitte den aktuellen Code am Eingang scannen").'
                  : 'The link opens the DEX app, recognizes you automatically via your login and checks you in — you see a green confirmation „Checked in!". You only ever check in yourself; the code cannot be used for other people. If something is off, the app shows the reason (e.g. „No registration found", „Check-in not open yet" or „Code expired — please scan the current code at the entrance").'}
              </>
            ),
            mockup: (
              <Callout variant="success" title={isDe ? 'Fertig in 2 Sekunden' : 'Done in 2 seconds'}>
                {isDe
                  ? 'Scannen → Link antippen → grüne Bestätigung. Kein Anstehen, kein Formular.'
                  : 'Scan → tap link → green confirmation. No queue, no form.'}
              </Callout>
            ),
          },
        ],
      },
    ],
  };
}
