/**
 * Rollen-Matrix - Übersicht aller Berechtigungen pro Rolle
 *
 * Nur für Admin zugänglich.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';

interface PermissionRow {
  category: string;
  feature: string;
  description: string;
  user: boolean | string;
  assistenz: boolean | string;
  /** v9.18+: Per-Event-Rolle "Co-Organizer" — gleiche Rechte wie Organizer am eigenen Event. */
  coorganizer: boolean | string;
  /** v9.16+: Globales Test-Team-Mitglied — sieht Entwurfs-Events, kann sich anmelden. */
  testteam: boolean | string;
  /** v6.19+: Per-Event-Rolle "Check-In-Team" (QR-Scanner) — nur Check-In-Tool. */
  checkin: boolean | string;
  organizer: boolean | string;
  admin: boolean | string;
}

const CATEGORY_COLORS: Record<string, { bg: string; border: string; text: string }> = {
  'Events ansehen':         { bg: 'rgba(59,130,246,0.08)',  border: '#3b82f6', text: '#1e40af' },
  'Event-Verwaltung':       { bg: 'rgba(134,188,37,0.10)',  border: '#86bc25', text: '#3f5f10' },
  'Registrierungen':        { bg: 'rgba(245,158,11,0.10)',  border: '#f59e0b', text: '#92400e' },
  'Teilnehmerverwaltung':   { bg: 'rgba(147,51,234,0.08)',  border: '#9333ea', text: '#6b21a8' },
  'Administration':         { bg: 'rgba(239,68,68,0.08)',   border: '#ef4444', text: '#991b1b' },
  'SharePoint':             { bg: 'rgba(107,114,128,0.10)', border: '#6b7280', text: '#374151' },
  'Profil':                 { bg: 'rgba(20,184,166,0.08)',  border: '#14b8a6', text: '#0f766e' },
};

const PERMISSIONS: PermissionRow[] = [
  // Events ansehen
  { category: 'Events ansehen', feature: 'Events des eigenen Standorts sehen',
    description: 'Events die per Location-/Audience-Filter für den eigenen Standort freigegeben sind in der Event-Liste sehen.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Events ansehen', feature: 'Alle Events sehen',
    description: 'Sieht auch Events, die per Location- oder Audience-Filter auf andere Standorte/Zielgruppen beschränkt sind.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Events ansehen', feature: 'Prozess-Übersicht (Flowcharts)',
    description: 'Zugriff auf die technischen Prozessdiagramme (Registrierungsflow, E-Mail-Pipeline, Outlook-Flow).',
    user: false, assistenz: false, coorganizer: true, testteam: false, checkin: false, organizer: true, admin: true },

  // Event-Verwaltung
  { category: 'Event-Verwaltung', feature: 'Events erstellen',
    description: 'Neues Event inkl. Subsite, Teilnehmerliste und Default-E-Mail-Templates anlegen.',
    user: false, assistenz: false, coorganizer: true, testteam: false, checkin: false, organizer: true, admin: true },
  { category: 'Event-Verwaltung', feature: 'Eigene Events bearbeiten',
    description: 'Metadaten (Titel, Zeiten, Ort, Filter, ...) von Events ändern, bei denen man in OrganizerEmail steht.',
    user: false, assistenz: false, coorganizer: true, testteam: false, checkin: false, organizer: true, admin: true },
  { category: 'Event-Verwaltung', feature: 'Alle Events bearbeiten',
    description: 'Auch fremde Events bearbeiten, für die man selbst nicht als Organizer hinterlegt ist.',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: true },
  { category: 'Event-Verwaltung', feature: 'Events löschen',
    description: 'Event unwiderruflich entfernen: Subsite, Teilnehmerliste und DEX_Events-Eintrag weg.',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Geteilte Kapazität konfigurieren',
    description: 'In Schritt 3 (Kapazität & Sichtbarkeit) zwei frei benannte Gruppen mit eigener Platzzahl anlegen — z.B. Vormittag/Nachmittag, VIP/Standard. Pro Gruppe kann zusätzlich eine eigene oder gemeinsame Warteliste eingestellt werden. Ersatz für die ehemalige B2Run-spezifische Durchstarter/Funstarter-Logik (seit v10.20).',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Pflichtfelder pro Split-Gruppe',
    description: 'In Schritt 4 (Felder) kann pro Custom-Field der Selector „Sichtbar für Teilnehmergruppe" auf „Beide", „Nur Gruppe A" oder „Nur Gruppe B" gesetzt werden — sichtbar nur wenn in Schritt 3 die geteilte Kapazität aktiv ist. Beispiel: Pflicht-Checkbox „Leistungsnachweis vorhanden" nur für die schnellere Lauf-Gruppe einblenden. Ersetzt seit v10.24 den hartkodierten Leistungsnachweis-Toggle.',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'B2Run-Events auf Standard-Schema migrieren',
    description: 'Pro B2Run-Legacy-Event ein Migrations-Button, der den B2Run-Type entfernt, b2run_*-Custom-Fields löscht und die Bezeichnungen "Durchstarter" / "Funstarter" als reguläre Gruppen-Labels persistiert. Anmeldungen, Wartelisten und Sub-Events bleiben unverändert. Nur für Admin sichtbar (seit v10.20).',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: true },
  { category: 'Event-Verwaltung', feature: 'Event-Bild hochladen (Item-Attachment)',
    description: 'Titelbild des Events hochladen/ersetzen — wird als Item-Attachment in DEX_Events gespeichert.',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Event-Dokumente hochladen',
    description: 'Zusatzdateien (PDFs, Agenda, Hotelinfo) anhängen — Teilnehmer sehen sie unter "Meine Events".',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Agenda / Transferzeiten / Quiz pflegen',
    description: 'Tages-Agenda, Bus-/Transferzeiten und Quiz-Fragen des Events anlegen und bearbeiten.',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'E-Mail-Templates pro Event anpassen',
    description: 'Subject/Heading/BodyHtml der Registrierungs-Mails für dieses eine Event überschreiben.',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'E-Mails pro Event deaktivieren',
    description: 'Automatische Bestätigungs-Mails (Anmeldung/Abmeldung/Warteliste) für dieses Event abschalten.',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Outlook-Einladungen pro Event deaktivieren',
    description: 'Automatische Outlook-Kalendereinträge für dieses Event abschalten.',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Sub-Events (Sessions) anlegen/editieren/löschen',
    description: 'Sub-Events (z.B. Trainingssessions bei B2Run) als eigene DEX_Events-Items mit gesetztem parentEventId anlegen. Jede Session hat eigene Teilnehmerliste, eigenen Outlook-Termin und eigene Mails. Gelöscht werden Child-Events kaskadierend mit dem Parent.',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Split-Kapazitäten (B2Run) aktivieren',
    description: 'In Schritt 3 (Kapazität) die Checkbox "Lauf-Event mit getrennten Starter-Kapazitäten" aktivieren. Ermöglicht separate Durchstarter/Funstarter-Zahlen mit eigenen Wartelisten und typ-bewusstem Nachrücken.',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },

  // Registrierungen
  { category: 'Registrierungen', feature: 'Selbst registrieren',
    description: 'Sich selbst für ein Event anmelden — solange die Anmeldefrist nicht abgelaufen ist. Bei Split-Kapazitäten (B2Run): Wunsch-Starter-Typ wählen. Ist er voll, bietet ein Dialog den Alt-Typ oder die Warteliste für den Wunsch-Typ.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Sub-Event-Sessions buchen',
    description: 'Bei der Anmeldung oder jederzeit nachträglich über "Meine Events" einzelne Sub-Events (Trainingssessions, optionale Workshops) buchen. Jede Session schreibt eine eigene Teilnehmerliste in ihrer Subsite und löst eine Bestätigungsmail + Outlook-Termin aus. Auch ohne Hauptevent-Anmeldung möglich (Sessions-Only-Modus).',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Sub-Event einzeln stornieren',
    description: 'Über "Meine Events" eine einzelne Sub-Event-Anmeldung zurückziehen, ohne dass die Hauptevent-Registrierung bzw. andere Sessions berührt werden. Versendet Cancellation-Mail und entfernt den Outlook-Termin der Session. Umgekehrt geht es genauso: Hauptevent abmelden, Sessions behalten — die Event-Karte bleibt mit Badge "Nur Sessions" sichtbar.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Auf Warteliste kommen',
    description: 'Wenn das Event (oder der Wunsch-Starter-Typ) voll ist, landet der User auf der Warteliste. PreferredStarterType wird für das typ-bewusste Nachrücken gespeichert.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Eigene Angaben bearbeiten',
    description: 'Eigene Registrierungsdaten (Custom Fields, T-Shirt-Größe, Notfallkontakt etc.) nachträglich ändern.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Eigene Registrierung stornieren',
    description: 'Eigene Anmeldung über "Meine Events" wieder zurückziehen; löst automatische Cancellation-Mail aus.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Für andere registrieren',
    description: 'Eine andere Person stellvertretend anmelden. Audit: RegisteredBy wird auf den eingeloggten User gesetzt.',
    user: false, assistenz: 'Nur Partner/Director ¹', coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Registrierungen', feature: 'Nach Anmeldefrist registrieren',
    description: 'Registrierungsformular auch nach Ablauf der RegistrationDeadline noch absenden.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Registrierungen', feature: 'Audit-Trail: RegisteredBy wird automatisch gesetzt',
    description: 'RegisteredByName/RegisteredByEmail werden bei jeder Anmeldung automatisch befüllt — unabhängig von der Rolle.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },

  // Teilnehmerverwaltung (Admin Center)
  { category: 'Teilnehmerverwaltung', feature: 'Teilnehmerliste sehen',
    description: 'Alle Teilnehmer des Events im Admin Center als Tabelle (mit Filter, Sortierung, Spaltenauswahl) sehen.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Teilnehmer suchen / sortieren',
    description: 'Freitextsuche + Spaltensortierung in der Admin-Teilnehmertabelle.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'QR-Code-Scanner / Check-In-Tool',
    description: 'Check-In-Bildschirm öffnen, QR-Codes scannen, Teilnehmer manuell ein-/auschecken. Check-In-Team (per Event) hat Zugriff aber NUR auf das Check-In-Tool — keine Bearbeitung, kein Mail-Versand, keine Teilnehmerliste.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: 'Eigene Events ²', organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Teilnehmer ein-/auschecken (manuell)',
    description: 'Check-in-Status eines Teilnehmers manuell setzen/zurücksetzen — ohne QR-Scanner.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: 'Eigene Events ²', organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Teilnehmer abmelden',
    description: 'Fremde Teilnehmer abmelden; löst Outlook-Ausladung und Cancellation-Mail aus. Seit v6.8: client-seitiger typ-bewusster Nachrück-Promote + Feedback-Toast "Nachgerückt: XY" (statt Warten auf Flow).',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Teilnehmer-Daten in der App bearbeiten',
    description: 'Edit-Button neben jedem Teilnehmer öffnet ein Modal zum direkten Anpassen von Vor-/Nachname und E-Mail-Adresse sowie aller Custom-Felder. Beim Ändern der E-Mail wird die Adresse gegen den Deloitte-Tenant validiert (nur @deloitte.de / @deloitte.com, Person muss existieren — externe Adressen werden abgewiesen, Tippfehler liefern eine klare Fehlermeldung). Profil-Felder (Phone/Department/Location/JobTitle) werden bei Mail-Wechsel automatisch nachgezogen. Jede Änderung landet im ChangeLog der Zeile (wer/wann/Vorher → Nachher) plus zentralem Audit-Log plus LastModifiedDate. Ersetzt das fehleranfällige direkte Editieren in der SharePoint-Liste.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Überbuchung prüfen & bereinigen',
    description: 'Findet pro Gruppe (bzw. gesamt) die zuletzt über Kapazität Angemeldeten (entstanden durch zeitgleiche Anmeldungen) und markiert sie — ohne Status-Änderung. Danach entscheidet der Bearbeiter pro Person: auf Warteliste setzen (optional mit Entschuldigungs-Mail im Deloitte-Layout und/oder Kalender-Abmeldung) oder Platz behalten (Erste(r) auf Warteliste bzw. bleibt angemeldet). Jede Korrektur landet im ChangeLog (war fälschlich angemeldet, Original-Registrierung). Danach werden TeilnehmerIDs automatisch neu vergeben (mit Fortschrittsanzeige). Seit v11.36; Organizer für eigene Events, weil es Teilnehmerverwaltung ist.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Getrennte Wartelisten sehen (B2Run-Split)',
    description: 'Bei Events mit aktivierter Split-Kapazität zeigt die Admin-Warteliste drei getrennte Tabellen (Durchstarter, Funstarter, "ohne Typ"). Zusätzliche Spalte "Startblock" in der Teilnehmer-Tabelle zeigt tatsächlichen StarterType + Wunsch-Typ in Klammern falls anders.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'QR-Codes versenden',
    description: 'Massen-Versand der persönlichen QR-Codes an alle bestätigten Teilnehmer.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'E-Mail-Adressen kopieren',
    description: 'Semikolon-separierte Liste aller Teilnehmer-Mails in die Zwischenablage kopieren.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Deep-Link auf Event-Admin kopieren',
    description: 'Link auf die Admin-Detail-Seite des aktuell offenen Events (?action=admin&event=<SP-ID>) in die Zwischenablage legen, um ihn an Co-Organizer oder Helfer weiterzugeben — sie landen nach Login direkt auf der gleichen Detail-Seite.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Massenmail an Teilnehmer',
    description: 'Freitext-Mail via RichText-Editor an alle Teilnehmer verschicken — im Deloitte-Template-Wrapper.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'IDs neu vergeben (Renummerierung)',
    description: 'TeilnehmerID sauber neu durchnummerieren (Aktive 1…N, Warteliste N+1…), inkl. Counter-Abgleich, mit %-Fortschrittsanzeige. Seit v11.36 auch für Organizer eigener Events (zählt als Teilnehmerverwaltung). Beim Öffnen eines Events mit kaputten IDs (Duplikate/fehlende IDs aus zeitgleichen Anmeldungen) erscheint automatisch ein Hinweis-Modal mit Korrigieren-Button — inkl. Warnung, dies nicht während laufender Anmeldewellen zu tun.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Spalten fixen (Schema reparieren)',
    description: 'Fehlende Basis-Spalten nachlegen + View-Reihenfolge korrigieren; entfernt B2Run/Quiz-Spalten auf Events, die sie nicht brauchen.',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Profile neu laden',
    description: 'Per Teilnehmer JobTitle/Standort/Department/Phone aus dem SP-User-Profil neu ziehen — inkl. UPN!=SMTP-Fallback.',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: true },

  // Administration
  { category: 'Administration', feature: 'Admin-Bereich öffnen',
    description: 'Zugriff auf das Admin Center mit Event-Liste und Management-Aktionen.',
    user: false, assistenz: false, coorganizer: true, testteam: false, checkin: false, organizer: true, admin: true },
  { category: 'Administration', feature: 'Rollen verwalten',
    description: 'Rollen von Usern (User / Organizer / Admin) in DEX_Roles hinzufügen, ändern, entfernen.',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: true },
  { category: 'Administration', feature: 'Test-Team verwalten (global)',
    description: 'Globale Liste von Personen pflegen, die alle Entwurfs-Events sehen + sich anmelden können — auch ohne Organizer/Admin-Rolle. Per People-Picker hinzufügen/entfernen, gespeichert auf dem _Config-Item der DEX_EmailTemplates-Liste.',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: true },
  { category: 'Administration', feature: 'Co-Organizer pro Event hinzufügen',
    description: 'Im Event-Wizard (Schritt 2) einen oder mehrere weitere Deloitte-User als Co-Organizer für dieses eine Event eintragen. Diese haben dann gleiche Rechte wie der Hauptorganizer am Event — können auch ihrerseits weitere Co-Organizer und Check-In-Personen hinzufügen.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Administration', feature: 'Check-In-Team pro Event hinzufügen',
    description: 'Im Event-Wizard (Schritt 2) Personen als reines Check-In-Team eintragen. Sie sehen das Event in ihrer Liste, dürfen NUR das Check-In-Tool nutzen — keine Bearbeitung, keine Teilnehmerliste, keine Mails.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Administration', feature: 'Event aktivieren / deaktivieren',
    description: 'Im Admin-Center per Toggle-Button zwischen "Active" (für alle berechtigten User sichtbar + buchbar) und "Under Construction" (nur Organizer/Admin/Test-Team sichtbar) wechseln. Schnellster Weg, ein Event zu pausieren oder live zu schalten.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Administration', feature: 'Counter zurücksetzen (TeilnehmerID-Recovery)',
    description: 'Setzt den DEX_TeilnehmerCounter auf den aktuellen Max-TID-Wert + repariert dabei die Visitors-Permissions auf der Counter-Liste. Recovery-Button für den (sehr seltenen) Fall, dass der Counter unter den Max-TID gefallen ist.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Administration', feature: 'QR-Code Auto-Send bei Anmeldung',
    description: 'Per Toggle pro Event: jede neue Anmeldung bekommt automatisch ihren QR-Code direkt mit der Bestätigungsmail (statt manuell Massen-Versand). Wird im "QR-Codes versenden"-Modal aktiviert.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Administration', feature: 'Onboarding-Mail an neuen Organizer/Admin senden',
    description: 'Nach Anlage einer neuen Organizer- oder Admin-Rolle bietet die App an, eine Begrüßungsmail im Deloitte-Layout zu verschicken (Links zur App, zum Handbuch, Kurzanleitung Test-Event). ebrenneisen@deloitte.de und nifelten@deloitte.de stehen automatisch im Cc.',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: true },
  { category: 'Administration', feature: 'Rollen-Matrix einsehen',
    description: 'Diese Übersichtsseite öffnen.',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: true },
  { category: 'Administration', feature: 'User suchen',
    description: 'Tenant-weite User-Suche (People-Picker) in den Admin-Settings nutzen.',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: true },
  { category: 'Administration', feature: 'Standort-Filter konfigurieren',
    description: 'Globale Standort-Liste pflegen, die in Event-Filtern (locationAudience) verwendet wird.',
    user: false, assistenz: false, coorganizer: true, testteam: false, checkin: false, organizer: true, admin: true },
  { category: 'Administration', feature: 'Zielgruppen-Filter konfigurieren',
    description: 'Globale Audience-Gruppen (z.B. "M&A", "SR&T") pflegen; werden in Event-Filtern genutzt.',
    user: false, assistenz: false, coorganizer: true, testteam: false, checkin: false, organizer: true, admin: true },
  { category: 'Administration', feature: 'Globale E-Mail-Templates bearbeiten',
    description: 'Standard-Templates in DEX_EmailTemplates (gilt für alle Events ohne eigene Overrides) anpassen.',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: true },

  // SharePoint (Visitors = DEALL, Owners = Admins)
  { category: 'SharePoint', feature: 'DEX_Events: Lesen',
    description: 'Leserechte auf die zentrale Event-Liste.',
    user: 'Visitors (Read)', assistenz: 'Visitors (Read)', coorganizer: 'Contribute', testteam: 'Visitors (Read)', checkin: false, organizer: 'Contribute', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_Events: Schreiben (inkl. Item-Attachments)',
    description: 'Schreibrechte auf DEX_Events inkl. Bild + Dokumenten-Attachments.',
    user: false, assistenz: false, coorganizer: 'Contribute', testteam: false, checkin: false, organizer: 'Contribute', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_Roles: Lesen',
    description: 'Leserechte auf die Rollen-Liste (für UI-Rollenerkennung).',
    user: false, assistenz: false, coorganizer: 'Read', testteam: false, checkin: false, organizer: 'Read', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_Roles: Schreiben',
    description: 'Schreibrechte auf DEX_Roles (Rollen-Zuweisung).',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_Emails: Queue (eigene)',
    description: 'Mail-Queue-Liste; mit Item-Level-Security: User sieht nur seine eigenen Einträge.',
    user: 'Contribute + ILS', assistenz: 'Contribute + ILS', coorganizer: 'Contribute + ILS', testteam: 'Contribute + ILS', checkin: false, organizer: 'Contribute + ILS', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_Outlook: Queue (eigene)',
    description: 'Outlook-Termin-Queue; mit Item-Level-Security.',
    user: 'Contribute + ILS', assistenz: 'Contribute + ILS', coorganizer: 'Contribute + ILS', testteam: 'Contribute + ILS', checkin: false, organizer: 'Contribute + ILS', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_IDReorder: Queue',
    description: 'Queue für TeilnehmerID-Renummerierungen — nur Admin, triggert Power Automate Batch.',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_EmailTemplates: Schreiben',
    description: 'Schreibrechte auf die globale Template-Liste.',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_Participants: eigene Einträge',
    description: 'Participant-Directory mit Item-Level-Security — User sieht/schreibt nur den eigenen Eintrag.',
    user: 'Contribute + ILS', assistenz: 'Contribute + ILS', coorganizer: 'Contribute + ILS', testteam: 'Contribute + ILS', checkin: false, organizer: 'Contribute + ILS', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'Event-Subsite',
    description: 'Zugriffsrechte auf die Subsite des Events (enthält die Teilnehmerliste).',
    user: 'Visitors (Read)', assistenz: 'Visitors (Read)', coorganizer: 'Full Control', testteam: 'Visitors (Read)', checkin: false, organizer: 'Full Control', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'Teilnehmerliste: eigener Eintrag',
    description: 'Lese-/Schreibrechte auf den eigenen Teilnehmer-Eintrag über Item-Level-Security.',
    user: 'Contribute + ILS', assistenz: 'Contribute + ILS', coorganizer: 'Full Control', testteam: 'Contribute + ILS', checkin: false, organizer: 'Full Control', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_TeilnehmerCounter (pro Subsite)',
    description: 'Atomarer Counter für TeilnehmerID-Vergabe (ein Item, NextValue-Feld). Visitors brauchen Contribute, damit der ETag-CAS-Inkrement bei der Anmeldung durchgeht — ohne Schreibrechte landet die TID auf null und die Anmeldung ist unvollständig.',
    user: 'Contribute', assistenz: 'Contribute', coorganizer: 'Full Control', testteam: 'Contribute', checkin: 'Contribute', organizer: 'Full Control', admin: 'Full Control' },
  { category: 'SharePoint', feature: 'DEX_ChangeLog: Audit-Log',
    description: 'Zentrales Audit-Log für Event- und Teilnehmer-Änderungen. Wird von der App geschrieben (registerForEvent, cancelRegistration, adminUpdateRegistration etc.). Lesezugriff für Admin zur Nachvollziehbarkeit.',
    user: false, assistenz: false, coorganizer: 'Append-only', testteam: false, checkin: false, organizer: 'Append-only', admin: 'Full Control' },

  // Profil
  { category: 'Profil', feature: 'Eigenes Profil ansehen',
    description: 'Eigene Profilseite mit Name, Rolle, Office, JobTitle öffnen.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Profil', feature: 'Settings-Seite öffnen',
    description: 'Persönliche Einstellungen (Sprache, Profilbild-Refresh) öffnen.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
];

function renderCell(value: boolean | string): React.ReactElement {
  if (value === true) {
    return <span style={{ color: '#22c55e', fontWeight: 700, fontSize: '1.1rem' }}>&#10003;</span>;
  }
  if (value === false) {
    return <span style={{ color: '#ef4444', fontWeight: 500, fontSize: '1.1rem' }}>&mdash;</span>;
  }
  return <span style={{ color: '#d97706', fontSize: '0.78rem', fontWeight: 600 }}>{value}</span>;
}

export default function RoleMatrixPage(): React.ReactElement {
  const { navigate } = useNavigation();

  const categories = Array.from(new Set(PERMISSIONS.map(p => p.category)));

  return (
    <div className="page-container">
      <div className="card" style={{ padding: 0, overflow: 'hidden' }}>
        <div style={{
          padding: '24px 28px 16px',
          borderBottom: '1px solid var(--dex-gray-200)',
          background: 'linear-gradient(135deg, rgba(134,188,37,0.08) 0%, rgba(59,130,246,0.05) 100%)',
        }}>
          <h2 style={{ margin: 0, fontSize: '1.2rem', fontWeight: 700 }}>Rollen-Matrix</h2>
          <p style={{ margin: '6px 0 0', color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>
            Übersicht aller Berechtigungen nach Rolle
          </p>
        </div>

        <div style={{ overflowX: 'auto', overflowY: 'auto', maxHeight: 'calc(100vh - 260px)' }}>
          <table style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, fontSize: '0.85rem' }}>
            <thead>
              <tr>
                <th style={{ ...thStyle, ...stickyHeaderCell, width: '32%', borderRight: '1px solid var(--dex-gray-300)' }}>Funktion</th>
                <th style={{ ...thStyle, ...stickyHeaderCell, textAlign: 'center', width: '9%', borderRight: '1px solid var(--dex-gray-300)' }}>
                  <span style={roleBadgeStyle('#3b82f6')}>User</span>
                </th>
                <th style={{ ...thStyle, ...stickyHeaderCell, textAlign: 'center', width: '10%', borderRight: '1px solid var(--dex-gray-300)' }}>
                  <span style={roleBadgeStyle('#9333ea')}>Assistenz</span>
                </th>
                <th style={{ ...thStyle, ...stickyHeaderCell, textAlign: 'center', width: '10%', borderRight: '1px solid var(--dex-gray-300)' }}>
                  <span style={roleBadgeStyle('#0ea5e9')}>Test-Team</span>
                </th>
                <th style={{ ...thStyle, ...stickyHeaderCell, textAlign: 'center', width: '10%', borderRight: '1px solid var(--dex-gray-300)' }}>
                  <span style={roleBadgeStyle('#14b8a6')}>Check-In</span>
                </th>
                <th style={{ ...thStyle, ...stickyHeaderCell, textAlign: 'center', width: '10%', borderRight: '1px solid var(--dex-gray-300)' }}>
                  <span style={roleBadgeStyle('#10b981')}>Co-Organizer</span>
                </th>
                <th style={{ ...thStyle, ...stickyHeaderCell, textAlign: 'center', width: '10%', borderRight: '1px solid var(--dex-gray-300)' }}>
                  <span style={roleBadgeStyle('#f59e0b')}>Organizer</span>
                </th>
                <th style={{ ...thStyle, ...stickyHeaderCell, textAlign: 'center', width: '9%' }}>
                  <span style={roleBadgeStyle('#86bc25')}>Admin</span>
                </th>
              </tr>
            </thead>
            <tbody>
              {categories.map(cat => {
                const rows = PERMISSIONS.filter(p => p.category === cat);
                const cc = CATEGORY_COLORS[cat] || { bg: 'transparent', border: 'var(--dex-gray-400)', text: 'var(--dex-gray-700)' };
                return (
                  <React.Fragment key={cat}>
                    <tr>
                      <td colSpan={8} style={{
                        padding: '14px 20px 10px 16px',
                        fontWeight: 700,
                        fontSize: '0.82rem',
                        textTransform: 'uppercase' as const,
                        letterSpacing: '0.08em',
                        color: cc.text,
                        background: cc.bg,
                        borderLeft: `4px solid ${cc.border}`,
                        borderTop: '1px solid var(--dex-gray-300)',
                        borderBottom: `2px solid ${cc.border}55`,
                      }}>
                        {cat}
                      </td>
                    </tr>
                    {rows.map((row, idx) => (
                      <tr key={row.feature} style={{
                        background: idx % 2 === 0 ? '#fff' : 'var(--dex-gray-50, #fafafa)',
                      }}>
                        <td style={{ ...dataCellStyle, borderRight: '1px solid var(--dex-gray-200)', padding: '12px 20px', color: 'var(--dex-gray-800)', lineHeight: 1.4 }}>
                          <div style={{ fontWeight: 600, marginBottom: 3 }}>{row.feature}</div>
                          <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>{row.description}</div>
                        </td>
                        <td style={{ ...dataCellStyle, borderRight: '1px solid var(--dex-gray-200)', padding: '12px 8px', textAlign: 'center', verticalAlign: 'middle' }}>{renderCell(row.user)}</td>
                        <td style={{ ...dataCellStyle, borderRight: '1px solid var(--dex-gray-200)', padding: '12px 8px', textAlign: 'center', verticalAlign: 'middle' }}>{renderCell(row.assistenz)}</td>
                        <td style={{ ...dataCellStyle, borderRight: '1px solid var(--dex-gray-200)', padding: '12px 8px', textAlign: 'center', verticalAlign: 'middle' }}>{renderCell(row.testteam)}</td>
                        <td style={{ ...dataCellStyle, borderRight: '1px solid var(--dex-gray-200)', padding: '12px 8px', textAlign: 'center', verticalAlign: 'middle' }}>{renderCell(row.checkin)}</td>
                        <td style={{ ...dataCellStyle, borderRight: '1px solid var(--dex-gray-200)', padding: '12px 8px', textAlign: 'center', verticalAlign: 'middle' }}>{renderCell(row.coorganizer)}</td>
                        <td style={{ ...dataCellStyle, borderRight: '1px solid var(--dex-gray-200)', padding: '12px 8px', textAlign: 'center', verticalAlign: 'middle' }}>{renderCell(row.organizer)}</td>
                        <td style={{ ...dataCellStyle, padding: '12px 8px', textAlign: 'center', verticalAlign: 'middle' }}>{renderCell(row.admin)}</td>
                      </tr>
                    ))}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>

        <div style={{
          padding: '16px 28px', borderTop: '1px solid var(--dex-gray-200)',
          display: 'flex', gap: 24, fontSize: '0.8rem', color: 'var(--dex-gray-500)',
          flexWrap: 'wrap',
        }}>
          <span><span style={{ color: '#22c55e', fontWeight: 700 }}>&#10003;</span> = Berechtigt</span>
          <span><span style={{ color: '#ef4444', fontWeight: 500 }}>&mdash;</span> = Kein Zugriff</span>
          <span><span style={{ color: '#d97706', fontWeight: 600 }}>Text</span> = Eingeschränkt</span>
        </div>
        <div style={{
          padding: '12px 28px 20px', borderTop: '1px dashed var(--dex-gray-200)',
          fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.6,
        }}>
          <div style={{ marginBottom: 6 }}>
            <strong>Rolle &quot;Assistenz&quot;</strong> ist keine DEX_Roles-Rolle, sondern wird automatisch aktiv, wenn der eingeloggte User im Azure-AD-Profil einen <em>JobTitle</em> hat, der &quot;Assistant&quot; enthält (deckt <em>&quot;Assistant&quot;</em> und <em>&quot;Senior Assistant&quot;</em> ab). Sonst verhält sich die Assistenz wie ein normaler User.
          </div>
          <div style={{ marginBottom: 6 }}>
            <strong>¹ Nur Partner/Director:</strong> Assistenz darf &quot;Für andere registrieren&quot; nutzen, aber nur für Personen mit JobTitle <strong>Partner</strong> oder <strong>Director</strong>. Die Assistenz muss sich selber für das Event auch anmelden können (d.h. die Anmeldefrist darf nicht abgelaufen sein).
          </div>
          <div>
            <strong>² Eigene Events:</strong> gilt nur wenn der User in <code>OrganizerEmail</code> des jeweiligen Events steht. Tenant-weiter Organizer-Status reicht nicht — Event A-Organizer können keine Admin-Aktionen für Event B ausführen. Admin darf global alles.
          </div>
        </div>
      </div>

      <div style={{ marginTop: 16, textAlign: 'center' }}>
        <button className="btn btn-secondary" onClick={() => navigate('settings')}>
          Zurück zu Settings
        </button>
      </div>
    </div>
  );
}

const thStyle: React.CSSProperties = {
  padding: '14px 16px', textAlign: 'left', fontWeight: 600, fontSize: '0.85rem', color: 'var(--dex-gray-700)',
};

// Sticky Kopfzeile: bleibt beim Scrollen sichtbar
const stickyHeaderCell: React.CSSProperties = {
  position: 'sticky' as const, top: 0, zIndex: 10,
  background: '#fff',
  borderBottom: '2px solid var(--dex-gray-400)',
};

// Zellen-Trenner: horizontale Zeilen-Linie sichtbarer
const dataCellStyle: React.CSSProperties = {
  borderBottom: '1px solid var(--dex-gray-200)',
};

function roleBadgeStyle(color: string): React.CSSProperties {
  return {
    display: 'inline-block', padding: '4px 12px', borderRadius: 14,
    background: `${color}22`, color: color, fontWeight: 700, fontSize: '0.82rem',
    letterSpacing: '0.02em',
  };
}
