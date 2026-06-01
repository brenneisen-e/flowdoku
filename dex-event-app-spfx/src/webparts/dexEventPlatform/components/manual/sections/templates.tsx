import * as React from 'react';
import { ManualSection } from '../types';
import { Callout } from '../ManualMockups';
import { AppPreview } from '../previews/AppPreview';
import Header from '../../Header';
import SettingsPage from '../../SettingsPage';

export function templatesSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'email-templates',
    title: isDe ? 'Globale E-Mail-Templates' : 'Global email templates',
    category: 'admin',
    description: isDe
      ? 'Die vier System-Mails (Anmeldung, Warteliste, Abmeldung, Nachrücken) zentral anpassen. Pro Event können sie überschrieben werden.'
      : 'Centrally customize the four system emails (Registration, Waitlist, Cancellation, Promotion). Each event can override them.',
    visibleFor: ['Admin'],
    perspectives: [
      {
        perspective: 'admin',
        steps: [
          {
            number: 1,
            title: isDe ? 'Templates öffnen' : 'Open templates',
            description: (
              <>
                {isDe
                  ? 'In den Einstellungen → "E-Mail-Templates". Du siehst alle vier System-Typen in DE und EN, jeweils mit Subject, Heading und HTML-Body. Klick auf "Bearbeiten" öffnet den HTML-Editor mit Live-Vorschau.'
                  : 'In Settings → "Email templates". You see all four system types in DE and EN, each with subject, heading and HTML body. Click "Edit" to open the HTML editor with live preview.'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Platzhalter nutzen' : 'Use placeholders',
            description: (
              <>
                <p style={{ margin: '0 0 6px 0' }}>
                  {isDe
                    ? 'In Templates funktionieren Platzhalter, die beim Versand pro Empfänger ersetzt werden:'
                    : 'Templates support placeholders that are replaced per recipient at send time:'}
                </p>
                <ul style={{ paddingLeft: 18, margin: '0 0 6px 0', lineHeight: 1.7 }}>
                  <li><code>{'{{Name}}'}</code> — {isDe ? 'Vorname des Teilnehmers' : 'attendee\'s first name'}</li>
                  <li><code>{'{{EventTitle}}'}</code> — {isDe ? 'Event-Titel' : 'event title'}</li>
                  <li><code>{'{{Organizer}}'}</code> — {isDe ? 'Liste der Organizer-Namen, sprachsensitiv formatiert' : 'list of organizer names, language-aware formatting'}</li>
                  <li><code>{'{{AppUrl}}'}</code> — {isDe ? 'Link zurück in die App' : 'link back to the app'}</li>
                  <li><code>{'{{WaitlistPosition}}'}</code> — {isDe ? 'nur in Warteliste-Mails: aktuelle Position' : 'waitlist mails only: current position'}</li>
                </ul>
                <p style={{ margin: 0 }}>
                  {isDe
                    ? 'Im Editor kannst du Platzhalter per Klick auf den entsprechenden Chip direkt an der Cursor-Position einfügen — du musst sie nicht abtippen.'
                    : 'In the editor, click a placeholder chip to insert it directly at the cursor position — no typing required.'}
                </p>
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Settings → Globale E-Mail-Templates (echte Ansicht)' : 'Settings → global email templates (real view)'}
                role="Admin"
                page="settings"
                width={1024}
                device="laptop"
              >
                <Header />
                <SettingsPage />
              </AppPreview>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Per-Event-Overrides verstehen' : 'How per-event overrides work',
            description: (
              <>
                <p style={{ margin: '0 0 6px 0' }}>
                  {isDe
                    ? 'Beim Erstellen eines Events kann der Organizer im Schritt "Kommunikation" einzelne System-Mails überschreiben (Subject + Heading + HTML-Body). Diese Überschreibungen gelten NUR für dieses eine Event und werden in DEX_Events.EmailTemplateOverrides als JSON gespeichert.'
                    : 'When creating an event, the organizer can override individual system mails in the "Communication" step (subject + heading + HTML body). These overrides apply ONLY to that specific event and are stored as JSON in DEX_Events.EmailTemplateOverrides.'}
                </p>
                <p style={{ margin: '0 0 6px 0' }}>
                  {isDe
                    ? 'Beim Versand wird pro Feld diese Reihenfolge geprüft:'
                    : 'On send, each field is resolved in this order:'}
                </p>
                <ol style={{ paddingLeft: 22, margin: '0 0 6px 0', lineHeight: 1.7 }}>
                  <li>{isDe ? 'Event-spezifischer Override (wenn gesetzt)' : 'Event-specific override (if set)'}</li>
                  <li>{isDe ? 'Globales SP-Template aus DEX_EmailTemplates' : 'Global SP template from DEX_EmailTemplates'}</li>
                  <li>{isDe ? 'Code-Default (Hardcoded Fallback im EmailTemplates.ts)' : 'Code default (hardcoded fallback in EmailTemplates.ts)'}</li>
                </ol>
                <p style={{ margin: 0 }}>
                  {isDe
                    ? 'Die headingColor (Brand-Farbe pro Mail-Typ) wird NIE überschrieben — sie kommt immer aus dem globalen Template.'
                    : 'The headingColor (brand color per mail type) is NEVER overridden — it always comes from the global template.'}
                </p>
              </>
            ),
            mockup: (
              <Callout variant="info" title={isDe ? 'Bug-Historie' : 'Bug history'}>
                {isDe
                  ? 'Vor v5.6 wurden Per-Event-Overrides zwar gespeichert, aber beim Versand nie ausgelesen — die App hat immer das globale Template benutzt. Seit v5.6 funktioniert die Override-Resolution korrekt (siehe applyEventTemplateOverride in EventContext.tsx).'
                  : 'Prior to v5.6, per-event overrides were saved but never consulted at send time — the app always used the global template. Since v5.6 the override resolution works correctly (see applyEventTemplateOverride in EventContext.tsx).'}
              </Callout>
            ),
          },
          {
            number: 4,
            title: isDe ? 'Live-Vorschau & DEX-Orb' : 'Live preview & DEX orb',
            description: (
              <>
                {isDe
                  ? 'Der HTML-Editor zeigt rechts eine Live-Vorschau im Deloitte-Mail-Wrapper. Das Deloitte-Logo (oben) und die DEX-Orb (unter dem Header) werden aus dem Cache (DEX_EmailTemplates._Config.LogoBase64 + DefaultImageBase64) geladen — gleiches Bild wie beim echten Versand. Wenn du einen frischen Tab benutzt und die Bilder fehlen, einmal Hard-Reload (Strg+Shift+R) löst das Cache-Loading aus.'
                  : 'The HTML editor shows a live preview in the Deloitte mail wrapper on the right. The Deloitte logo (top) and DEX orb (below the header) are loaded from cache (DEX_EmailTemplates._Config.LogoBase64 + DefaultImageBase64) — same images as in real sends. If you opened a fresh tab and images are missing, a hard-reload (Ctrl+Shift+R) triggers cache loading.'}
              </>
            ),
          },
          {
            number: 5,
            title: isDe ? 'Editor-Toolbar' : 'Editor toolbar',
            description: (
              <>
                {isDe
                  ? 'Im HTML-Editor stehen Fett, Kursiv, Unterstrichen, Schriftgröße (das Dropdown zeigt die Größe des markierten Texts wie in Word), Brand-Farben plus freie Farbwahl über Farb-Picker oder Hex-Code, Listen (Bullet + nummeriert) und ein Format-Reset zur Verfügung. Enter erzeugt einen <br> (E-Mail-konform), Shift+Enter einen neuen Absatz. Über dem Body lassen sich Überschrift UND Unter-Überschrift separat formatieren (Größe, Farbe, fett, kursiv).'
                  : 'The HTML editor offers Bold, Italic, Underline, font size (the dropdown shows the selected text\'s size like in Word), brand colors plus free color choice via color picker or hex code, lists (bullet + numbered), and a format reset. Enter inserts a <br> (email-friendly), Shift+Enter starts a new paragraph. Above the body, the heading AND subheading can each be styled separately (size, color, bold, italic).'}
              </>
            ),
            tip: isDe
              ? 'Vermeide externe CSS-Klassen oder ausgefallene HTML-Strukturen — viele Mailclients (vor allem Outlook) ignorieren sie. Inline-Styles und einfache Tags (<p>, <strong>, <em>, <a>, <ul>, <li>) sind am robustesten.'
              : 'Avoid external CSS classes or fancy HTML structures — many mail clients (especially Outlook) ignore them. Inline styles and simple tags (<p>, <strong>, <em>, <a>, <ul>, <li>) are the most robust.',
          },
          {
            number: 6,
            title: isDe ? 'Logos & Default-Bilder' : 'Logos & default images',
            description: (
              <>
                {isDe
                  ? 'Im gleichen Bereich verwaltest du die globalen Bilder: Deloitte-Logo für den Mail-Header (LogoBase64), DEX-Orb für den Hero-Bereich der Mail (DefaultImageBase64) und Default-Event-Bilder (falls Organizer keines hochladen).'
                  : 'In the same area you manage the global images: Deloitte logo for the mail header (LogoBase64), DEX orb for the mail hero area (DefaultImageBase64), and default event images (used when organizers don\'t upload one).'}
              </>
            ),
            mockup: <Callout variant="info">{isDe ? 'Alle Bilder werden als Base64 in DEX_EmailTemplates._Config gespeichert — keine externen URL-Abhängigkeiten. Der Cache wird einmal pro Session geladen (loadLogosAsBase64).' : 'All images are stored as Base64 in DEX_EmailTemplates._Config — no external URL dependencies. The cache is loaded once per session (loadLogosAsBase64).'}</Callout>,
          },
          {
            number: 7,
            title: isDe ? 'Outlook-Body — Sonderfall' : 'Outlook body — special case',
            description: (
              <>
                {isDe
                  ? 'Der Outlook-Termin-Body funktioniert ähnlich wie System-Mails (HTML-Editor, Toolbar, Variablen), aber mit anderen Variablen (Address, Location, StartDate, EndDate) — diese werden bereits beim Speichern in der App ersetzt (nicht erst beim Versand), weil der Outlook-Body für ALLE Teilnehmer identisch ist. Per-Empfaenger-Variablen (Name, AppUrl) ergeben hier keinen Sinn.'
                  : 'The Outlook event body works similarly to system mails (HTML editor, toolbar, variables) but uses different variables (Address, Location, StartDate, EndDate) — these are resolved at save time (not at send time) because the Outlook body is identical for all attendees. Per-recipient variables (Name, AppUrl) don\'t make sense here.'}
              </>
            ),
            warning: isDe
              ? 'Der Outlook-Body wird vom Power-Automate-Flow DEX_Outlook_Einladungen erst seit 2026-04-17 beim "Outlook event aktualisieren" mit gepatcht. Falls du ältere Events updatest und der Body nicht erscheint, prüfe ob beide Flows (DEX_CreateOutlookEvent + DEX_Outlook_Einladungen) den aktuellen Stand haben.'
              : 'The Outlook body is patched on "Update Outlook event" only since 2026-04-17 in the DEX_Outlook_Einladungen flow. If you update older events and the body doesn\'t appear, verify both flows (DEX_CreateOutlookEvent + DEX_Outlook_Einladungen) are on the latest version.',
          },
        ],
      },
    ],
  };
}
