/**
 * Rollen-Matrix - Übersicht aller Berechtigungen pro Rolle
 *
 * Nur für Admin zugänglich.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { getCachedLogoBase64, getCachedOrbBase64 } from '../services/EmailTemplates';
import { DELOITTE_LOGO_BLACK } from '../data/brandLogos';

// v24.13: Das gecachte Deloitte-Logo ist WEISS (für dunkle Mail-Header). Fürs
// PDF (weißer Grund) färben wir genau dieses offizielle Logo per Canvas auf
// Schwarz um — Form/Proportionen bleiben das Original, nur die Farbe ändert sich.
async function recolorLogoBlack(dataUrl: string): Promise<string | null> {
  if (!dataUrl) return null;
  return new Promise<string | null>((resolve) => {
    const img = new Image();
    img.onload = (): void => {
      try {
        const c = document.createElement('canvas');
        c.width = img.naturalWidth || 400; c.height = img.naturalHeight || 100;
        const ctx = c.getContext('2d');
        if (!ctx) { resolve(null); return; }
        ctx.drawImage(img, 0, 0, c.width, c.height);
        ctx.globalCompositeOperation = 'source-in';
        ctx.fillStyle = '#000000';
        ctx.fillRect(0, 0, c.width, c.height);
        resolve(c.toDataURL('image/png'));
      } catch { resolve(null); }
    };
    img.onerror = (): void => resolve(null);
    img.src = dataUrl;
  });
}

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
  { category: 'Event-Verwaltung', feature: 'Entwürfe / Events ohne fremde Anmeldungen löschen',
    description: 'Events, die nie aktiv waren bzw. keine Anmeldungen über das Organizer-Team hinaus haben (z. B. Entwürfe), können mit einer einfachen Ja-Bestätigung gelöscht werden. Abgelaufene Entwürfe werden dem Organizer dafür auf der Startseite vorgeschlagen.',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Ehemals aktive Events löschen (mit echten Anmeldungen)',
    description: 'Events, die einmal aktiv waren und Anmeldungen über das Organizer-Team hinaus hatten, dürfen NUR Admins löschen — und erst ein Jahr nach dem Event-Ende (die Teilnehmerliste muss ein Jahr aufbewahrt werden).',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: 'Erst nach 1 Jahr' },
  { category: 'Event-Verwaltung', feature: 'Abgelaufenes Event archivieren (aus eigener Übersicht ausblenden)',
    description: 'Nach dem Event kann der Organizer es archivieren — dann verschwindet es aus SEINER Event-Übersicht (im Hintergrund bleibt es mit allen Daten erhalten). Andere sehen es weiter.',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Administration', feature: 'Hintergrund-Daten archivieren (Mails/Kalender/Protokolle)',
    description: 'Zeilen abgelaufener Events aus den Arbeitslisten ins Archiv (DEX_Archive) verschieben, damit diese schlank bleiben. Die Teilnehmerliste selbst wird NICHT archiviert (1 Jahr Aufbewahrung).',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: true },
  { category: 'Administration', feature: 'Alte Archiv-Einträge endgültig löschen',
    description: 'Archiv-Einträge, die älter als 1 Monat (nach Event-Ablauf) sind, endgültig aus DEX_Archive entfernen. Hinweis dazu beim App-Start für Admins.',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: true },
  { category: 'Event-Verwaltung', feature: 'Geteilte Kapazität konfigurieren',
    description: 'In Schritt 3 (Kapazität & Sichtbarkeit) zwei frei benannte Gruppen mit eigener Platzzahl anlegen — z.B. Vormittag/Nachmittag, VIP/Standard. Pro Gruppe kann zusätzlich eine eigene oder gemeinsame Warteliste eingestellt werden. Ersatz für die ehemalige B2Run-spezifische Durchstarter/Funstarter-Logik (seit v10.20).',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Pflichtfelder pro Split-Gruppe',
    description: 'In Schritt 5 (Felder) kann pro Custom-Field der Selector „Sichtbar für Teilnehmergruppe" auf „Beide", „Nur Gruppe A" oder „Nur Gruppe B" gesetzt werden — sichtbar nur wenn in Schritt 3 die geteilte Kapazität aktiv ist. Beispiel: Pflicht-Checkbox „Leistungsnachweis vorhanden" nur für die schnellere Lauf-Gruppe einblenden. Ersetzt seit v10.24 den hartkodierten Leistungsnachweis-Toggle.',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Anrede-Toggle pro Event',
    description: 'In Schritt 5 (Felder) per Checkbox „Anrede abfragen?" entscheiden, ob das Registrierungsformular ein Anrede-Dropdown (Frau / Herr / Divers / Keine Angabe) anzeigt. Default aus — viele Events brauchen die Anrede nicht und ersparen Teilnehmern das Feld. Eingeführt mit v11.80.',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Sicherheitshinweis vor Anmeldung konfigurieren',
    description: 'In Schritt 5 (Felder), Section ganz unten, einen Bestätigungs-Dialog aktivieren, der nach „Anmelden" und vor der eigentlichen Anmeldung erscheint. Zwei Modi: „Auswahl-Übersicht" (listet Haupt-Event + gewählte Sub-Events; der Teilnehmer kann vor dem Absenden einzelne Punkte ab-/zuwählen) oder „Eigener Hinweistext" (frei formulierter Hinweis mit Pflicht-Bestätigung). Default aus. Eingeführt mit v18.75.',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Event-Verwaltung', feature: 'Dokument-Upload-Feld anlegen (v19.0)',
    description: 'In Schritt 5 (Felder) den Feldtyp „Dokument (PDF/Bild-Upload)" wählen, um Teilnehmer um einen Datei-Upload zu bitten (PDF/JPG/PNG, max. 10 MB). Die Datei wird an die Teilnehmer-Zeile angehängt (kein Spaltenwert) und ist im Admin Center pro Teilnehmer abruf- und löschbar. Als Pflichtfeld konfigurierbar. Eingeführt mit v19.0.',
    user: false, assistenz: false, coorganizer: 'Eigene', testteam: false, checkin: false, organizer: 'Eigene', admin: true },
  { category: 'Registrierungen', feature: 'Dokument hochladen (v19.0)',
    description: 'Bei Events mit Dokument-Feld eine Datei (PDF/JPG/PNG, max. 10 MB) hochladen — direkt bei der Anmeldung oder später über „Meine Events" ergänzen/ersetzen/löschen. Die Datei hängt an der eigenen Teilnehmer-Zeile; der Organizer sieht sie im Admin Center. Eingeführt mit v19.0.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Event-Verwaltung', feature: 'Team-Anmeldung konfigurieren',
    description: 'In Schritt 4 (Team-Anmeldung) sechs Settings konfigurieren — Basis: Toggle „Team-Anmeldung erlauben", Team-Größe (Default 4, Min 2, Max 20), Toggle „Team-Namen abfragen". Beitritts-Modus (v11.81): Radio „Nur komplette Teams" vs. „Auch Teil-Teams erlaubt" (Default komplett), Checkbox „Unvollständige Teams öffentlich für Beitritt sichtbar" (Default aus), Checkbox „Beitritt erfordert Bestätigung durch Team-Kapitän" (Default aus, nur aktivierbar wenn Sichtbarkeit an). Eingeführt mit v11.80 + erweitert v11.81 als reine Konfiguration — die tatsächliche Multi-Person-Anmelde-Logik (Form, automatische Mails, Outlook-Einladungen, Slot-Beitritt, Approve-Queue) folgt mit v11.82+.',
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
  { category: 'Registrierungen', feature: 'Nicht-Teilnahme melden (Absage ohne Anmeldung)',
    description: 'Über die Anmelde-Seite per Button „Ich nehme nicht teil" proaktiv zurückmelden, dass man NICHT teilnimmt — ohne sich vorher anzumelden. Es wird kein Sitzplatz belegt; der Eintrag landet im Admin-Center unter „Abmeldungen" und ist dort als „Absage (nicht angemeldet)" gekennzeichnet (separat von regulären Stornierungen). Wer es sich anders überlegt, kann sich danach jederzeit normal anmelden.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Sub-Event-Sessions buchen',
    description: 'Bei der Anmeldung oder jederzeit nachträglich über "Meine Events" einzelne Sub-Events (Trainingssessions, optionale Workshops) buchen. Jede Session schreibt eine eigene Teilnehmerliste in ihrer Subsite und löst eine Bestätigungsmail + Outlook-Termin aus. Auch ohne Hauptevent-Anmeldung möglich (Sessions-Only-Modus).',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Sub-Event einzeln stornieren',
    description: 'Über "Meine Events" eine einzelne Sub-Event-Anmeldung zurückziehen, ohne dass die Hauptevent-Registrierung bzw. andere Sessions berührt werden. Versendet Cancellation-Mail und entfernt den Outlook-Termin der Session. Umgekehrt geht es genauso: Hauptevent abmelden, Sessions behalten — die Event-Karte bleibt mit Badge "Nur Sessions" sichtbar.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Team-Anmeldung für sich + Team durchführen',
    description: 'Bei Events mit aktivierter Team-Anmeldung (Schritt 4) zusätzlich N-1 weitere Personen mit einer einzigen Submission anmelden. Pflicht: vorab Zustimmung jedes Teammitglieds einholen + im Formular per Checkbox bestätigen. Jedes Mitglied bekommt automatisch Anmeldebestätigung, Outlook-Termin und sieht das Event in „Meine Events". Reicht die Restkapazität nicht für das ganze Team, geht das gesamte Team gemeinsam auf die Warteliste (keine Teil-Aktivierung). Eingeführt mit v11.82.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Team-Mitglied nachträglich hinzufügen (nur als Team-Lead)',
    description: 'Auf „Meine Events" sieht der Team-Lead bei freien Slots einen Button „+ Mitglied hinzufügen". Modal mit People-Picker, orange Pflicht-Hinweisbox zur Zustimmung des neuen Mitglieds und Pflicht-Bestätigungs-Checkbox. Beim Klick wird die Person sofort in das Team eingetragen — Anmeldebestätigung + Outlook-Einladung gehen direkt raus, die anderen Team-Mitglieder bekommen eine Info-Mail „X ist eurem Team beigetreten". Doppel-Anmelde-Schutz prüft, dass die Person nicht schon beim Event registriert ist. Eingeführt mit v11.83.',
    user: 'Als Team-Lead', assistenz: 'Als Team-Lead', coorganizer: 'Als Team-Lead', testteam: 'Als Team-Lead', checkin: false, organizer: 'Als Team-Lead', admin: 'Als Team-Lead' },
  { category: 'Registrierungen', feature: 'Offenem Team beitreten (Beitritt aus Anmelde-Seite)',
    description: 'Wenn der Organizer „Offene Slots öffentlich sichtbar" aktiviert hat, sehen alle Teilnehmer auf der Event-Anmeldeseite eine Box (seit v18.73 UNTER der Team-Anmelde-Karte) mit allen unvollständigen Teams (Team-Name, Belegung — Mitgliedernamen aus Datenschutzgründen ausgeblendet). Seit v18.73 wird ein Team per „Vormerken" nur ausgewählt — die eigentliche Anmeldung (inkl. der event-spezifischen Pflichtfelder) erfolgt erst über den „Anmelden"-Button; bei Approval-Events wird stattdessen eine Anfrage in der Approve-Queue angelegt. Eingeführt mit v11.83, überarbeitet v18.73.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Team-Beitritts-Anfragen bearbeiten (nur als Team-Lead)',
    description: 'Wenn der Organizer „Beitritt erfordert Bestätigung durch Team-Kapitän" aktiviert hat, sieht der Team-Lead auf „Meine Events" einen orange Block „Beitritts-Anfragen (N)" mit den offenen Anfragen. Pro Anfrage kann der Lead „Bestätigen" (Aufnahme erfolgt direkt inkl. Mails) oder „Ablehnen" (kurze Absage-Mail an den Anfragenden) klicken. Anfragen-Status landet in der globalen DEX_TeamJoinRequests-Liste. Eingeführt mit v11.83.',
    user: 'Als Team-Lead', assistenz: 'Als Team-Lead', coorganizer: 'Als Team-Lead', testteam: 'Als Team-Lead', checkin: false, organizer: 'Als Team-Lead', admin: 'Als Team-Lead' },
  { category: 'Registrierungen', feature: 'Auto-Promote nach Team-Lead-Abmeldung',
    description: 'Meldet sich der Team-Lead über „Meine Events" ab, wird automatisch das früheste verbleibende Mitglied (kleinste TeilnehmerID, sonst früheste Registration-Date) zum neuen Team-Lead promotet. Alle verbleibenden Mitglieder bekommen eine Info-Mail, der neue Lead sieht in seiner Mail einen extra Hinweis. Wenn keine Mitglieder mehr übrig sind, löst sich das Team auf. Eingeführt mit v11.83.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Team-Mitglieder als Lead aus dem Team abmelden',
    description: 'Auf „Meine Events" sieht der Team-Lead im Team-Badge der eigenen Event-Karte einen Button „Team bearbeiten" neben „+ Mitglied hinzufügen". Klick öffnet das Modal „Team verwalten" mit allen Mitgliedern als Karten (Profilfoto, Name, E-Mail, Standort, Lead-Badge). Pro aktivem Mitglied außer dem Lead selbst gibt es einen roten Trash-Button — Klick fordert eine zweite Bestätigung an, dann wird die Person stellvertretend vom Event abgemeldet. Der Lead-Audit-Eintrag (CancelledByName/Email) zeigt den Lead, nicht die abgemeldete Person. Abmelde-Bestätigungs-Mail, Outlook-Termin-Absage, ID-Neuvergabe und Info-Mail an die verbleibenden Mitglieder laufen automatisch. Der Lead selbst sieht keinen Trash-Button für sich — sein Self-Cancel läuft über den normalen Abmelden-Button mit Auto-Promote. Eingeführt mit v11.86.',
    user: 'Als Team-Lead', assistenz: 'Als Team-Lead', coorganizer: 'Als Team-Lead', testteam: 'Als Team-Lead', checkin: false, organizer: 'Als Team-Lead', admin: 'Als Team-Lead' },
  { category: 'Registrierungen', feature: 'Auf Warteliste kommen',
    description: 'Wenn das Event (oder der Wunsch-Starter-Typ) voll ist, landet der User auf der Warteliste. PreferredStarterType wird für das typ-bewusste Nachrücken gespeichert.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Eigene Angaben bearbeiten',
    description: 'Eigene Registrierungsdaten (Custom Fields, T-Shirt-Größe, Notfallkontakt etc.) nachträglich ändern.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Eigene Registrierung stornieren',
    description: 'Eigene Anmeldung über "Meine Events" wieder zurückziehen; löst automatische Cancellation-Mail aus. Seit v22.22 nur bis zum Event-Ende möglich — bei vergangenen Events ist die Selbst-Abmeldung gesperrt (auch über den Abmelde-Link aus Mails); Organizer/Admins können im Organizer Center weiterhin abmelden, dann aber still (ohne Mail, Outlook-Absage und Nachrücken).',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: false, organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Für andere registrieren',
    description: 'Eine andere Person stellvertretend anmelden. Pflicht-Checkbox zur Zustimmung der Person; diese Bestätigung wird seit v18.74 zusätzlich als Nachweis in der Teilnehmerliste gespeichert (Spalte „ProxyConsent", mit Name + Datum des Akteurs). Audit: RegisteredBy wird auf den eingeloggten User gesetzt.',
    user: false, assistenz: 'Nur Partner/Director ¹', coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Registrierungen', feature: 'Assistenz bei eigener Anmeldung verknüpfen',
    description: 'Admins und Directoren können bei der eigenen Event-Anmeldung eine Assistenz angeben (Personen-Suche). Die Assistenz kommt auf CC der Bestätigung und sieht die Anmeldung als INFO in ihrer „Assistenz"-Kachel; die anmeldende Person verwaltet die Anmeldung weiterhin selbst. Die Assistenz kann (in Kürze) eine Änderung/Abmeldung anfordern. Erscheint nicht, wenn das Event bereits ein Assistenz-CC-Feld hat. Reine App-Lösung über DEX_AssistantAccess, ohne Flow. Eingeführt mit v24.41.',
    user: false, assistenz: false, coorganizer: 'Nur als Director', testteam: false, checkin: false, organizer: 'Nur als Director', admin: true },
  { category: 'Registrierungen', feature: 'Assistenz: stellvertretende Anmeldungen verwalten',
    description: 'Über die Startseiten-Kachel „Assistenz" alle Anmeldungen verwalten, die man SELBST stellvertretend für andere durchgeführt hat: Angaben einsehen und anpassen, die Person bei Sub-Events an- und abmelden sowie ganz abmelden. Gezeigt werden nur die eigenen Fremd-Anmeldungen (Filter: RegisteredBy = ich, Teilnehmer ≠ ich) — eine Anmeldung, die die Person selbst vorgenommen hat, ist hier nicht sichtbar. Die Kachel erscheint nur, wenn man tatsächlich solche Anmeldungen hat; Admins sehen sie immer und finden darin ebenfalls ihre eigenen Fremd-Anmeldungen. Eingeführt mit v24.36.',
    user: 'Wenn vorhanden', assistenz: true, coorganizer: true, testteam: 'Wenn vorhanden', checkin: 'Wenn vorhanden', organizer: true, admin: true },
  { category: 'Registrierungen', feature: 'Externe Person (außerhalb Deloitte) anmelden',
    description: 'Im „Für andere registrieren"-Modus per Checkbox „Person außerhalb Deloitte" eine externe Person erfassen — People-Picker aus, Vorname/Nachname/E-Mail frei eintragbar. Zustimmung muss SCHRIFTLICH vorliegen (Pflicht-Checkbox + ProxyConsent-Nachweis). Tippfehler-Prüfung der E-Mail + Kontroll-Dialog vor dem Absenden. Die Bestätigungs-Mail geht direkt an die externe Person mit den Organizern auf CC; ein Outlook-Termin wird an externe Adressen NICHT versendet. Nur Organizer/Admin (nicht Assistenz). Eingeführt mit v18.74.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
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
  { category: 'Event-Verwaltung', feature: 'People-Picker-Feld → Person auf CC der An-/Abmelde-Mail (v18.41)',
    description: 'Pro People-Picker-Custom-Field (Typ „Person", z.B. „Assistenz") kann der Organizer einstellen, dass die ausgewählte Person die Anmelde- und Abmelde-Mail des Teilnehmers automatisch in Kopie (CC) bekommt. Betrifft nur die E-Mails, nicht den Outlook-Termin.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Self-Check-in: QR/Live-Anzeige bereitstellen (v18.33, v20.2: immer möglich)',
    description: 'Self-Check-in ist für jedes Event immer möglich — der Organizer entscheidet selbst, ob er es nutzt, indem er den QR-Code bereitstellt (oder eben nicht). QR-PDF herunterladen bzw. rotierende Live-Anzeige öffnen: direkt auf der Check-in-Seite, im Admin Center, aus dem QR-Versand-Modal oder (ab 5 Tage vor dem Event bzw. nach QR-Versand) über die QR-Kachel unter dem Event-Bild im Event-Detail. Das Check-in-Zeitfenster (Von/Bis, verhindert verfrühte UND nachträgliche Check-ins) wird im Kachel-Modal eingestellt; der frühere Wizard-Schalter ist entfallen. Teilnehmer scannen den event-spezifischen QR-Code mit der Handy-Kamera und checken sich selbst ein.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: 'Wenn vom Organizer schon geöffnet', organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Sich selbst per QR-Code einchecken (v18.33)',
    description: 'Wenn der Organizer Self-Check-in aktiviert hat: am Veranstaltungstag den event-spezifischen QR-Code mit der normalen Handy-Kamera scannen und sich selbst als anwesend markieren. Jeder checkt ausschließlich sich selbst ein (Login-gebunden); der Code lässt sich nicht für andere verwenden.',
    user: true, assistenz: true, coorganizer: true, testteam: true, checkin: true, organizer: true, admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Teilnehmer abmelden',
    description: 'Fremde Teilnehmer abmelden; löst Outlook-Ausladung und Cancellation-Mail aus. Seit v6.8: client-seitiger typ-bewusster Nachrück-Promote + Feedback-Toast "Nachgerückt: XY" (statt Warten auf Flow).',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Teilnehmer-Daten in der App bearbeiten',
    description: 'Edit-Button neben jedem Teilnehmer öffnet ein Modal zum direkten Anpassen von Vor-/Nachname und E-Mail-Adresse sowie aller Custom-Felder. Beim Ändern der E-Mail wird die Adresse gegen den Deloitte-Tenant validiert (nur @deloitte.de / @deloitte.com, Person muss existieren — externe Adressen werden abgewiesen, Tippfehler liefern eine klare Fehlermeldung). Profil-Felder (Phone/Department/Location/JobTitle) werden bei Mail-Wechsel automatisch nachgezogen. Jede Änderung landet im ChangeLog der Zeile (wer/wann/Vorher → Nachher) plus zentralem Audit-Log plus LastModifiedDate. Ersetzt das fehleranfällige direkte Editieren in der SharePoint-Liste.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Abgemeldete Registrierungen löschen (v19.28)',
    description: 'In der „Abmeldungen"-Liste pro Zeile ein Löschen-Button (nach Sicherheits-Bestätigung). Die abgemeldete Registrierung wird endgültig aus der Teilnehmerliste entfernt (kein Recycle-Bin) — z.B. um Test-Anmeldungen aus der Übersicht zu räumen. Die Löschung landet im Audit-Log (RegistrationDeleted, wer/wann).',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Übergreifende (Klammer-)Felder pro Teilnehmer bearbeiten (v19.30)',
    description: 'In der konsolidierten Ansicht eines Events mit Sub-Events lassen sich die übergreifenden Hauptevent-Custom-Felder pro Teilnehmer direkt bearbeiten („Felder"-Button + Modal). Speichert in die Hauptevent-Registrierung und protokolliert die Änderung (Vorher → Nachher je Feld) im Audit-Log.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Teilnehmer von einzelnen Sub-Events / allen abmelden (v19.30)',
    description: 'In der konsolidierten Ansicht öffnet ein „Abmelden"-Button ein Modal mit allen Sub-Events, in denen die Person registriert ist (Checkbox-Liste + „Alle auswählen"). Pro gewähltem Sub-Event laufen Abmelde-Mail, Outlook-Ausladung, Nachrücken und ID-Neuvergabe automatisch — die jeweiligen Mail-/Outlook-Schalter des Sub-Events werden respektiert. Jede Abmeldung landet im Audit-Log (RegistrationCancelled).',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Audit-Log / Änderungsprotokoll pro Event ansehen (v19.30)',
    description: 'Über die Aktion „Audit-Log / Änderungsprotokoll" das Änderungsprotokoll geöffnet, vorgefiltert auf das aktuelle Event. Zeigt pro Eintrag Zeitpunkt, Akteur (wer), Aktionstyp, betroffene Person und bei Daten-Änderungen das Vorher → Nachher je Feld. Erfasst Anmeldungen, Abmeldungen, Löschungen, stellvertretende Anmeldungen und Feld-Änderungen.',
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
  { category: 'Teilnehmerverwaltung', feature: 'Einladungsmail mit Anmelde-Link versenden',
    description: 'Mail mit Anmelde-Link an dich selbst (zum Weiterleiten an externe/interne Verteiler) oder direkt an den auf dem Event hinterlegten Mailverteiler verschicken — Default-Text vorbefüllt, im RichText-Editor frei editierbar, im Deloitte-Template-Wrapper.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'IDs neu vergeben (Renummerierung)',
    description: 'TeilnehmerID sauber neu durchnummerieren (Aktive 1…N, Warteliste N+1…), inkl. Counter-Abgleich, mit %-Fortschrittsanzeige. Seit v11.36 auch für Organizer eigener Events (zählt als Teilnehmerverwaltung). Die IDs sind durch den Flow ohnehin durchlaufend; öffnet man ein Event kurz nach einer Abmeldung, erscheint ein Hinweis-Modal, dass die automatische Korrektur (Nachrücken + ID-Neuvergabe) evtl. noch läuft und man ein paar Minuten warten soll, statt parallel manuell zu korrigieren.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Personen zu offenem Team hinzufügen (Admin Center)',
    description: 'Im Admin Center, oberhalb der Teilnehmer-Tabelle, sieht der Organizer/Admin pro Event mit aktivierter Team-Anmeldung eine Sektion „Teams (N)" mit allen aktiven Teams. Pro Team ein Button „Person hinzufügen" (nur bei freien Slots), öffnet das gleiche orange Hinweisbox-Modal wie in „Meine Events" — People-Picker, Pflicht-Bestätigungs-Checkbox. Neue Person wird sofort angemeldet, bekommt Bestätigungs-Mail + Outlook-Termin, die anderen Mitglieder bekommen eine Info-Mail. Doppel-Anmelde-Schutz prüft, dass die Person nicht schon beim Event registriert ist. Eingeführt mit v11.84.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Team-Lead-Rolle übergeben (Admin Center)',
    description: 'Pro Team im Admin-Center-Teams-Block ein Button „Lead-Rolle übergeben". Öffnet ein Dropdown mit allen anderen aktiven Mitgliedern. Auswahl setzt die alte Lead-Zeile auf TeamLead=false und die neue auf TeamLead=true — danach kann der neue Lead in „Meine Events" Mitglieder hinzufügen und Beitritts-Anfragen entscheiden. Alle Mitglieder bekommen automatisch eine Info-Mail im Deloitte-Layout, der neue Lead extra mit dem Hinweis auf seine erweiterten Rechte. Eingeführt mit v11.84.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Event-Verwaltung', feature: 'Team-Begriff frei benennen + Teilnehmer-Sperre (v23.0)',
    description: 'In Schritt 4 (Team-Anmeldung) kann der Begriff „Team" frei umbenannt werden — getrennt für Einzahl und Mehrzahl (z.B. „Break-Out Session"/„Break-Out Sessions"). Der gewählte Begriff erscheint überall in der App (Anmeldeformular, Organizer Center, „Meine Events"-Badge inkl. „<Begriff>-Lead"). Zusätzlich ein Schalter „Teilnehmer dürfen keine neuen Teams/Gruppen erstellen": ist er aktiv, verschwindet die „Ich melde mich + mein Team an"-Karte auf der Anmeldeseite — die Aufteilung übernimmt dann der Organizer selbst per Drag & Drop. Eingeführt mit v23.0.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Personen per Drag & Drop zu Teams zuordnen (Admin Center, v23.0)',
    description: 'In der Teams-Sektion des Organizer Centers zieht der Organizer/Admin Personen mit der Maus zwischen den Teams/Break-Out-Sessions und der „ohne Team"-Box hin und her (HTML5 Drag & Drop, Ziel-Feld färbt sich grün). Loslassen ordnet die Person dem Team zu oder löst die Zuordnung. War der Gezogene ein Team-Lead und bleiben Mitglieder übrig, rückt das früheste verbleibende Mitglied als neuer Lead nach. Jede Verschiebung wird im Änderungsprotokoll festgehalten. Eingeführt mit v23.0.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Per-Team-Info-Mail senden (z.B. Einwahllink, v23.0)',
    description: 'Sobald aktive Teams existieren, gibt es in der Teams-Sektion den Button „Mail an <Teams>". Im Modal trägt der Organizer/Admin Betreff + Mail-Text (Platzhalter {{Vorname}}, {{Name}}, {{TeamName}}, {{EventTitle}}, {{TeamInfo}}) sowie PRO Team ein eigenes Info-Feld ein — typischerweise einen dedizierten Microsoft-Teams-Einwahllink je Break-Out-Session. Beim Versand bekommt jedes aktive Mitglied eine eigene Mail im Deloitte-Layout, in der {{TeamInfo}} durch die Info seines Teams ersetzt ist (Links werden klickbar). Respektiert den E-Mail-Schalter aus Schritt 6. Eingeführt mit v23.0.',
    user: false, assistenz: false, coorganizer: 'Eigene Events ²', testteam: false, checkin: false, organizer: 'Eigene Events ²', admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Spalten fixen (Schema reparieren)',
    description: 'Fehlende Basis-Spalten nachlegen + View-Reihenfolge korrigieren; entfernt B2Run/Quiz-Spalten auf Events, die sie nicht brauchen.',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: true },
  { category: 'Administration', feature: 'Archivierung abgelaufener Event-Zeilen (v22)',
    description: 'Beim App-Start sieht der Admin rechts auf der Startseite eine Archivierungs-Box, sobald Zeilen abgelaufener Events anstehen. Per Klick werden alle betroffenen Zeilen aus den Arbeitslisten (DEX_Emails, DEX_Outlook, DEX_IDReorder, DEX_ChangeLog, DEX_AccessFix) in die Liste DEX_Archive verschoben — mit Fortschrittsanzeige. Das Archiv ist ausschließlich für Admins einsehbar; die Arbeitslisten bleiben dadurch schlank.',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: true },
  { category: 'Teilnehmerverwaltung', feature: 'Fremd-Anmeldungen: Zugriff reparieren (alle aktiven Events, v20.6)',
    description: 'Geht alle Teilnehmerlisten aktiver Events (inkl. Sub-Events) durch und prüft zwei Dinge: (1) jede Liste steht auf „nur eigene Elemente" — niemand kann fremde Anmeldedaten sehen; falls nicht, wird die Sicherheit direkt repariert. (2) Bei Anmeldungen, die jemand FÜR eine andere Person gemacht hat (Organizer/Admin/Assistenz), bekommt die angemeldete Person Zugriff auf ihre eigene Anmeldung — sie sieht sie danach in „Meine Events" und kann sich selbst abmelden. Externe Personen ohne Deloitte-Login können nicht berücksichtigt werden.',
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
  { category: 'Administration', feature: 'Event live schalten / auf Entwurf setzen',
    description: 'Im Admin-Center per Toggle-Button zwischen "Live" (für alle berechtigten User sichtbar + buchbar) und "Entwurf" (nur Organizer/Admin/Test-Team sichtbar) wechseln. Schnellster Weg, ein Event zu pausieren oder live zu schalten. Alternativ kann der Entwurf-Status auch beim Bearbeiten im Event-Wizard (Schritt 1) gesetzt werden.',
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
  { category: 'Administration', feature: 'Demo-Modus: als User testen (v12.7)',
    description: 'Im Header-User-Menü unter „Rollenverwaltung" einen beliebigen Deloitte-User + Standort wählen — die App agiert dann so als wäre man dieser User (Rolle „User", Standortfilter aktiv). Beendet wird die Impersonation über einen orangen Sticky-Banner ganz oben. Nützlich um Sicht und Berechtigungen eines Users zu prüfen ohne sich selbst auszuloggen.',
    user: false, assistenz: false, coorganizer: false, testteam: false, checkin: false, organizer: false, admin: true },
  { category: 'Administration', feature: 'Default-Mail-Templates re-seeden (v12.11+)',
    description: 'In den Settings-Einstellungen mit einem Klick alle Standard-Mail-Vorlagen in DEX_EmailTemplates (Anmeldung, Warteliste, Abmeldung, Nachrücken, Team-Mails, Zimmerpartner, Gruppen-Wechsel, Überbuchungs-Entschuldigung etc.) auf die im Code hinterlegten Default-Texte zurücksetzen. Eigene Anpassungen werden überschrieben. Notwendig nach App-Updates, die die Standard-Texte verbessert haben.',
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
  // v22.21: Geführtes Tutorial (Onboarding-Tour) auf der Landing Page.
  { category: 'Profil', feature: 'Geführtes Tutorial starten',
    description: 'Interaktive Tour durch die App (Landing Page, grüne Sprechblase). Alle sehen die Teilnehmer-Tour; wer Events anlegen kann oder Organizer/Co-Organizer eines Events ist, bekommt zusätzlich die Organizer-Tour zur Auswahl.',
    user: 'User-Tour', assistenz: 'User-Tour', coorganizer: 'User- + Organizer-Tour', testteam: 'User-Tour', checkin: 'User-Tour', organizer: 'User- + Organizer-Tour', admin: 'User- + Organizer-Tour' },
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

// v24.10: Rollen-Matrix als PDF (Querformat A4) — Deloitte-Logo oben links,
// DEX-App-Logo (Orb) oben rechts, danach die komplette Berechtigungstabelle.
async function downloadRoleMatrixPdf(): Promise<void> {
  const { jsPDF } = await import('jspdf');
  const doc = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'landscape' });
  const pageW = 297, pageH = 210, margin = 12;
  const cols = ['User', 'Assistenz', 'Test-Team', 'Check-In', 'Co-Org.', 'Organizer', 'Admin'];
  const featW = 96;
  const roleW = (pageW - 2 * margin - featW) / cols.length;
  const logoR = getCachedOrbBase64();
  // Schwarzes Deloitte-Logo oben links: bevorzugt das offizielle Repo-Logo,
  // sonst das gecachte weiße Logo auf Schwarz umgefärbt, sonst Text-Fallback.
  const logoLBlack = DELOITTE_LOGO_BLACK || (await recolorLogoBlack(getCachedLogoBase64())) || '';

  const drawChrome = (): number => {
    // Offizielles (schwarzes) Deloitte-Logo oben links.
    let drewLogo = false;
    if (logoLBlack) {
      try {
        const p = doc.getImageProperties(logoLBlack);
        const h = 11; const w = (p.width / p.height) * h;
        doc.addImage(logoLBlack, 'PNG', margin, 9, w, h);
        drewLogo = true;
      } catch { drewLogo = false; }
    }
    if (!drewLogo) {
      // Text-Fallback, falls kein Logo verfügbar ist.
      doc.setFont('helvetica', 'bold'); doc.setFontSize(19); doc.setTextColor(0, 0, 0);
      doc.text('Deloitte', margin, 18);
      const dWidth = doc.getTextWidth('Deloitte');
      doc.setTextColor(134, 188, 37); doc.text('.', margin + dWidth + 0.6, 18);
    }
    // DEX-App-Logo (Orb, farbig) oben rechts — auf Weiß gut sichtbar.
    try {
      if (logoR) { const p = doc.getImageProperties(logoR); const h = 13; const w = (p.width / p.height) * h; doc.addImage(logoR, 'PNG', pageW - margin - w, 8, w, h); }
    } catch { /* Logo optional */ }
    doc.setFont('helvetica', 'bold'); doc.setFontSize(15); doc.setTextColor(30);
    doc.text('Rollenmatrix — Berechtigungen', pageW / 2, 15, { align: 'center' });
    doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(120);
    doc.text(`DEX Event Experience Platform · Stand: ${new Date().toLocaleDateString('de-DE')}`, pageW / 2, 20, { align: 'center' });
    // Spaltenkopf
    const yy = 26;
    doc.setFillColor(134, 188, 37); doc.rect(margin, yy, pageW - 2 * margin, 8, 'F');
    doc.setTextColor(255); doc.setFont('helvetica', 'bold'); doc.setFontSize(8);
    doc.text('Funktion', margin + 2, yy + 5.3);
    cols.forEach((c, i) => doc.text(c, margin + featW + i * roleW + roleW / 2, yy + 5.3, { align: 'center' }));
    return yy + 8;
  };

  let y = drawChrome();
  let lastCat = '';
  const cell = (v: boolean | string): { t: string; c: [number, number, number] } =>
    v === true ? { t: 'Ja', c: [34, 140, 30] } : v === false ? { t: '–', c: [175, 175, 175] } : { t: String(v), c: [180, 95, 0] };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  for (const row of (PERMISSIONS as any[])) {
    doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5);
    const featLines = doc.splitTextToSize(row.feature, featW - 4) as string[];
    const vals = [row.user, row.assistenz, row.testteam, row.checkin, row.coorganizer, row.organizer, row.admin].map(cell);
    const cellLines = vals.map(v => doc.splitTextToSize(v.t, roleW - 3) as string[]);
    let maxLines = featLines.length;
    cellLines.forEach(cl => { if (cl.length > maxLines) maxLines = cl.length; });
    const rowH = Math.max(6, maxLines * 3.1 + 2.6);
    const catH = row.category !== lastCat ? 6 : 0;
    if (y + catH + rowH > pageH - margin) { doc.addPage(); y = drawChrome(); lastCat = ''; }
    if (row.category !== lastCat) {
      lastCat = row.category;
      doc.setFillColor(238, 240, 233); doc.rect(margin, y, pageW - 2 * margin, 6, 'F');
      doc.setTextColor(74, 124, 31); doc.setFont('helvetica', 'bold'); doc.setFontSize(8);
      doc.text(row.category, margin + 2, y + 4.2);
      y += 6;
    }
    doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(45);
    doc.text(featLines, margin + 2, y + 3.6);
    vals.forEach((v, i) => {
      doc.setTextColor(v.c[0], v.c[1], v.c[2]);
      doc.text(cellLines[i], margin + featW + i * roleW + roleW / 2, y + 3.6, { align: 'center' });
    });
    doc.setDrawColor(228); doc.line(margin, y + rowH, pageW - margin, y + rowH);
    y += rowH;
  }
  doc.save('DEX-Rollenmatrix.pdf');
}

export default function RoleMatrixPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const [pdfBusy, setPdfBusy] = React.useState(false);
  // v13.0: Admin-Guard hinzugefügt — laut CLAUDE.md-Rollenmatrix ist
  // "Rollen-Matrix einsehen" Admin-only. Vorher fehlte der Schutz —
  // jeder User (auch Demo-impersoniert) konnte die Seite öffnen. Wir
  // nutzen originalIsAdmin damit der Admin-im-Demo-Modus seine eigene
  // Matrix weiterhin testen kann.
  const { originalIsAdmin } = useRoles();
  React.useEffect(() => {
    if (!originalIsAdmin) navigate('start');
  }, [originalIsAdmin, navigate]);
  if (!originalIsAdmin) {
    return (
      <div className="page-container">
        <div className="card" style={{ padding: 24 }}>
          <p>Diese Seite ist nur für Administratoren zugänglich.</p>
        </div>
      </div>
    );
  }

  const categories = Array.from(new Set(PERMISSIONS.map(p => p.category)));

  return (
    <div className="page-container">
      <div className="card" style={{ padding: 0, overflow: 'hidden' }}>
        <div style={{
          padding: '24px 28px 16px',
          borderBottom: '1px solid var(--dex-gray-200)',
          background: 'linear-gradient(135deg, rgba(134,188,37,0.08) 0%, rgba(59,130,246,0.05) 100%)',
        }}>
          <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
            <div>
              <h2 style={{ margin: 0, fontSize: '1.2rem', fontWeight: 700 }}>Rollenmatrix</h2>
              <p style={{ margin: '6px 0 0', color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>
                Übersicht aller Berechtigungen nach Rolle
              </p>
            </div>
            {/* v24.10: PDF-Download (Querformat, mit Deloitte- + DEX-App-Logo). */}
            <button
              className="btn btn-secondary"
              style={{ fontSize: '0.82rem', padding: '8px 16px', flexShrink: 0 }}
              disabled={pdfBusy}
              onClick={() => { setPdfBusy(true); downloadRoleMatrixPdf().catch(() => { /* */ }).then(() => setPdfBusy(false)); }}
            >
              {pdfBusy ? 'PDF wird erstellt…' : 'Als PDF herunterladen'}
            </button>
          </div>
          {/* v18.5: Power User ist KEINE eigene Rolle, sondern ein Zusatz-Flag
              auf einem Organizer/Admin — daher keine eigene Matrix-Spalte. */}
          <div style={{
            marginTop: 12, padding: '10px 14px', borderRadius: 8,
            background: '#fff4e5', border: '1px solid #f0b67a',
            fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.55,
          }}>
            <span style={{
              display: 'inline-block', padding: '2px 10px', borderRadius: 12,
              background: '#fff', color: '#b35a00', fontWeight: 700, fontSize: '0.78rem', marginRight: 6,
            }}>★ Power User</span>
            ist <strong>keine eigene Rolle</strong>, sondern ein <strong>Zusatz-Status auf einem
            Organizer</strong> (oder Admin): dieselbe Person, ein Eintrag — gleichzeitig Organizer
            UND Power User. Power User kennen sich besonders gut aus und werden auf der
            Event-Erstellungs-Seite mit Name und Foto als <strong>Hilfe-Ansprechpartner</strong>
            angezeigt. Setzen/Entfernen per Stern-Button in der Rollenverwaltung — die
            Berechtigungen bleiben die eines Organizers.
          </div>
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
