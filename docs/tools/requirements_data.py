# -*- coding: utf-8 -*-
"""
Bewertung der System Security Requirements für DEX — einzige Quelle der Wahrheit.

Aus dieser Datei werden BEIDE Ergebnisse erzeugt:
  * docs/downloads/DEX-Requirements-kommentiert.xlsx  (Formular für den Review)
  * docs/security-requirements.html                   (verlinkte Matrix in der Doku)

Zwei Fassungen von Hand zu pflegen hiesse, dass sie irgendwann verschieden
antworten — genau der Fehler, den die Doku an anderer Stelle beschreibt.

Klassifikation (bewusst nur fünf Werte):
  platform  Erfüllt durch Microsoft 365 / den Deloitte-Tenant. DEX erbt es und
            fügt nichts hinzu. Nachweis führt der Plattformbetrieb.
  dex       Erfüllt durch DEX selbst, im Code oder im Betriebsablauf belegbar.
  na        Nicht anwendbar — mit Begründung, warum die Anforderung auf diese
            Bauart nicht zutrifft.
  open      Offen: DEX kann es nicht belegen, die Bestätigung muss aus dem
            Betrieb oder von GTS kommen.
  flagged   Abweichung oder Restrisiko, das bewusst offengelegt wird und eine
            Entscheidung braucht.

Grundsatz: Lieber „offen“ als eine Behauptung, die im Review nicht hält.
"""

# Anker in docs/architektur.html — hierhin verlinkt die Matrix je Anforderung.
CHAPTERS = {
    'k1': 'Das Systembild',
    'k2': 'Wo der Code läuft',
    'k3': 'Die Datenhaltung',
    'k4': 'Eine Subsite je Event',
    'k5': 'Die Zugriffsschicht im Code',
    'k6': 'Warteschlangen statt Aufrufe',
    'k7': 'Die sieben Flows',
    'k8': 'Ablauf: Anmeldung',
    'k9': 'Ablauf: Event anlegen',
    'k10': 'Ablauf: Abmelden und Nachrücken',
    'k11': 'Ablauf: Absage aus Outlook',
    'k12': 'Wo die Wahrheit liegt',
    'k13': 'Schutz und Aufbewahrung',
    'k15': 'Deployment und Rollback',
    'k16': 'Sicherung und Wiederherstellung',
    'k17': 'Sicherheitsbetrachtung',
}

# Stand der Auswertung — wird in beide Ergebnisse geschrieben.
APP_VERSION = '29.11.0'
ASSESSED_ON = '2026-08-12'
AUDIT_NOTE = ('npm audit vom 2026-08-12: 62 Befunde im Produktions-Abhängigkeitsbaum '
              '(2 kritisch, 22 hoch, 38 mittel) bei 2331 Paketen gesamt. Ein Teil davon '
              'sind Build-Werkzeuge, die im Bundle nicht ausgeliefert werden.')

REQS = [
    # ------------------------------------------------------------------ GAC
    dict(
        id='GAC-01', chapter='k7', cls='flagged',
        met='DEX bringt keine eigenen privilegierten Konten mit. Im Umfang liegen zwei: die '
            'Websitesammlungs-Administratoren der SharePoint-Site (tenantverwaltet) und die '
            'Verbindungs-Identität der sieben Power-Automate-Flows, die auf den Event-Subsites '
            'Vollzugriff braucht.',
        evidence='Architektur, Kapitel 7 „Die sieben Flows“; Auskunft des Betriebs vom 2026-08-17.',
        open='Befund mit bekanntem Abhilfeplan: Die Flow-Verbindungen laufen derzeit unter einem '
             'PERSÖNLICHEN Benutzerkonto, nicht unter einem Dienstkonto. Damit ist kein Vaulting im '
             'PAMS möglich — ein persönliches Konto gehört dort nicht hinein. Zwei Wirkungen: Die '
             'Anforderung ist nicht erfüllt, und der Betrieb der Plattform hängt an einer Person; '
             'scheidet sie aus, brechen alle Mails und Kalendereinträge ab. Geplant ist die Umstellung '
             'auf einen Service Principal. Damit wäre die Anforderung erfüllbar und GAC-02 sowie '
             'GAC-15 gleich mit. Bitte Zieltermin festhalten.'),
    dict(
        id='GAC-02', chapter='k7', cls='flagged',
        met='Die Flow-Verbindungen gehören der Identität, unter der sie angelegt wurden; die Freigabe '
            'an weitere Personen steuert Power Platform.',
        evidence='Architektur, Kapitel 7; Auskunft des Betriebs vom 2026-08-17.',
        open='Derzeit ist die Verbindungs-Identität ein persönliches Konto (siehe GAC-01). Formal ist '
             'die Anforderung damit erfüllt — die Zugangsdaten liegen ausschliesslich bei ihrem '
             'Eigentümer —, aber aus dem falschen Grund: Es gibt keinen autorisierten Vertreter, weil '
             'es keine Delegation gibt. Mit dem geplanten Service Principal wird daraus eine bewusst '
             'verwaltete Berechtigung mit benennbaren Verantwortlichen.'),
    dict(
        id='GAC-03', chapter='k2', cls='platform',
        met='DEX hat keine eigene Authentifizierung. Anmeldung, Sperrung nach Fehlversuchen und '
            'Benachrichtigung liegen vollständig bei Entra ID des Tenants.',
        evidence='Architektur, Kapitel 2 „Wo der Code läuft“ — die App handelt ausschliesslich mit '
                 'der bestehenden SharePoint-Sitzung des Nutzers.'),
    dict(
        id='GAC-04', chapter='k13', cls='dex',
        met='Auf Tenant-Ebene greift der reguläre Austrittsprozess. Auf Anwendungsebene führt DEX '
            'zusätzlich einen eigenen Abgleich: Ein Lauf prüft die hinterlegten Konten gegen Entra ID, '
            'meldet deaktivierte Konten an die Organizer und kann sie automatisch abmelden. '
            'Rollen in DEX_Roles werden dabei sichtbar.',
        evidence='Architektur, Kapitel 13 „Schutz und Aufbewahrung“; Funktionen scanInactiveAccounts, '
                 'notifyOrganizerOfInactive und autoDeregisterInactive im EventService.',
        open='Kadenz des Laufs ist manuell ausgelöst, nicht terminiert.'),
    dict(
        id='GAC-05', chapter='k13', cls='dex',
        met='Konten sind in DEX in drei Stufen kategorisiert: Teilnehmer (nur eigene Zeilen), '
            'Organizer je Event (Vollzugriff auf die eigene Event-Subsite) und Admin (Plattformweit). '
            'Dazu kommen Sonderrollen mit engem Umfang: Co-Organizer, Check-in-/QR-Team, Assistenz. '
            'Die Zuordnung steht in DEX_Roles beziehungsweise je Event am Event selbst.',
        evidence='Architektur, Kapitel 13; Handbuch-Artikel „Rollen und Rechte“ in der App.',
        open='Turnus der Rechteprüfung (zweimal jährlich laut Vorgabe) ist organisatorisch '
             'festzulegen — die App liefert die Liste, prüft aber nicht selbst.'),
    dict(
        id='GAC-09', chapter='k2', cls='platform',
        met='Kein eigener Anmeldevorgang in DEX. Sperrung und Entsperrung von Benutzerkonten regelt Entra ID.',
        evidence='Architektur, Kapitel 2.'),
    dict(
        id='GAC-10', chapter='k2', cls='platform',
        met='DEX unterhält keine eigene Sitzung, kein eigenes Token und kein eigenes Cookie. Die Seite '
            'lebt in der SharePoint-Sitzung; deren Gültigkeit und Leerlauf-Abbruch gelten unverändert.',
        evidence='Architektur, Kapitel 2 und 5 — jeder Aufruf läuft über den SPHttpClient der Seite.'),
    dict(
        id='GAC-11', chapter='k2', cls='platform',
        met='Abmelden erfolgt über die Microsoft-365-Kopfzeile der Seite; die Bestätigung zeigt Microsoft. '
            'Ein eigener Logout in DEX gäbe es nicht abzumelden.',
        evidence='Architektur, Kapitel 2.'),
    dict(
        id='GAC-13', chapter='k13', cls='dex',
        met='Rollenbasierte Zugriffssteuerung ist tragendes Element und wirkt auf zwei Ebenen. '
            'Fachlich: DEX_Roles (Admin, Organizer) entscheidet, wer Events anlegen und das Organizer '
            'Center öffnen darf; je Event kommen Organizer, Co-Organizer, Check-in-Team und Assistenz dazu. '
            'Technisch durchgesetzt: Die Teilnehmerlisten laufen mit Zeilen-Sicherheit („nur eigene Elemente“), '
            'geprüft von SharePoint am Zeilen-Autor — nicht im Browser. Wer in der Oberfläche etwas nicht '
            'sehen darf, bekommt es auch bei manipuliertem Client nicht.',
        evidence='Architektur, Kapitel 4 „Eine Subsite je Event“ und Kapitel 13; Handbuch „Rollen und Rechte“.'),
    dict(
        id='GAC-14', chapter='k13', cls='flagged',
        met='Der Zugang zur Anwendung erfolgt über die angemeldete Microsoft-365-Sitzung; der Hinweis auf '
            'Überwachung und zulässige Nutzung wird beim Tenant-Login gezeigt. DEX selbst blendet vor der '
            'Nutzung keinen eigenen Sicherheitshinweis ein; es gibt Nutzungsbedingungen als eigene Seite '
            'in der Anwendung.',
        evidence='Architektur, Kapitel 13.',
        open='Abweichung: Die drei geforderten Aussagen (Überwachung, Einwilligung durch Nutzung, '
             'Strafbarkeit unbefugter Nutzung) stehen nicht als Hinweis vor dem Zugang in DEX. '
             'Entscheidung nötig: Reicht der Tenant-Hinweis, oder soll DEX ihn wiederholen? '
             'Aufwand gering — ein Einblendtext vor dem ersten Öffnen.'),
    dict(
        id='GAC-15', chapter='k7', cls='flagged',
        met='Die Anforderung greift erst, wenn es ein nicht-menschliches Konto gibt.',
        evidence='Architektur, Kapitel 7; Auskunft des Betriebs vom 2026-08-17.',
        open='Heute nicht anwendbar, weil die Flows unter einem persönlichen Konto laufen (GAC-01) — '
             'und dem kann die interaktive Anmeldung nicht entzogen werden, es ist ja ein Mensch. Mit '
             'der Umstellung auf einen Service Principal löst sich die Anforderung von selbst: Ein '
             'Service Principal hat gar keine interaktive Anmeldung. Der Punkt ist also nicht separat '
             'zu bearbeiten, sondern erledigt sich mit GAC-01.'),
    # ------------------------------------------------------------------ GAT
    dict(
        id='GAT-01', chapter='k1', cls='dex',
        met='Dokumentation liegt in drei Ebenen vor und deckt die genannten Zielgruppen ab. '
            'Endnutzer und Organizer: Handbuch in der Anwendung, nach Perspektiven getrennt '
            '(Teilnehmer, Organizer, Admin), mit geführter Tour und durchsuchbaren Artikeln. '
            'Administratoren und Betrieb: Systemarchitektur, Flow-Übersicht, Flow-Definitionen. '
            'Entwicklung: Build- und Release-Ablauf, Arbeitsanweisungen im Repository, lückenlose '
            'Release Notes fachlich und technisch. Neue Organizer erhalten automatisch eine '
            'Einführungsmail.',
        evidence='Dokumentationsseite (Management Summary, Anwendungsbeschreibung, Flow-Übersicht, '
                 'Systemarchitektur, Release Notes) und Handbuch in der Anwendung.'),
    # ------------------------------------------------------------------ GAU
    dict(
        id='GAU-01', chapter='k13', cls='dex',
        met='Die Protokollierung teilt sich auf. Anmeldevorgänge, Fehlversuche, Sperrungen, '
            'Passwortänderungen, Abmeldungen sowie Quelle und Ziel des Zugriffs (Punkte 2 bis 7 und 10) '
            'protokolliert Microsoft 365 im einheitlichen Überwachungsprotokoll des Tenants — DEX '
            'erzeugt diese Ereignisse nicht selbst. Fachliche Vorgänge protokolliert DEX in DEX_ChangeLog: '
            'wer wann welche Aktion auf welchem Event ausgelöst hat, mit Ergebnis (Punkte 1, 8, 9, 11 bis 14). '
            'Erfasst werden unter anderem Anmeldung, Abmeldung, Rollenänderung, Löschung einer '
            'Teilnehmerliste und stellvertretende Anmeldungen.',
        evidence='Architektur, Kapitel 3 „Die Datenhaltung“ (Abschnitt Protokolle) und Kapitel 13.',
        open='IP-Adresse, Port und Protokoll erfasst DEX bewusst nicht selbst; sie stehen im '
             'Überwachungsprotokoll des Tenants.'),
    dict(
        id='GAU-02', chapter='k13', cls='platform',
        met='DEX betreibt keine eigene Protokollinfrastruktur. Fachliche Einträge liegen in einer '
            'SharePoint-Liste derselben Site, Plattformereignisse im Überwachungsprotokoll von Microsoft 365.',
        evidence='Architektur, Kapitel 3.'),
    dict(
        id='GAU-03', chapter='k13', cls='flagged',
        met='Zeitstempel werden als ISO-8601 in UTC gespeichert (Beispiel 2026-08-12T09:14:22Z) — das '
            'native Format der SharePoint-Datumsspalte.',
        evidence='Architektur, Kapitel 3.',
        open='Abweichung vom geforderten Format yyyymmdd-hhmmss. Begründung: Das Format ist durch '
             'SharePoint vorgegeben; ISO 8601 ist eindeutig, sortierbar und enthält dieselbe Information '
             'inklusive Zeitzone. Eine Umformatierung wäre nur in einer Auswertung möglich, nicht in der '
             'Speicherung. Bitte um Bewertung, ob das akzeptiert wird.'),
    dict(
        id='GAU-04', chapter='k13', cls='platform',
        met='Alle Zeitstempel entstehen serverseitig in SharePoint Online beziehungsweise in Power '
            'Automate und sind UTC. Die App rechnet nur für die Anzeige in Berliner Zeit um.',
        evidence='Architektur, Kapitel 3.'),
    dict(
        id='GAU-05', chapter='k13', cls='flagged',
        met='DEX_ChangeLog ist eine SharePoint-Liste; Lesen ist auf Admins beschränkt, Löschen ebenfalls. '
            'Änderungen an Listenelementen zeichnet SharePoint in der Versionshistorie und im '
            'Überwachungsprotokoll auf.',
        evidence='Architektur, Kapitel 3 und 13.',
        open='Restrisiko: Die Liste ist nicht technisch schreibgeschützt (append-only). Ein Admin könnte '
             'Einträge ändern; nachweisbar bliebe das über Versionshistorie und Tenant-Überwachung. '
             'Entscheidung nötig, ob das genügt oder ob die Einträge zusätzlich nach GEMS gespiegelt '
             'werden sollen.'),
    dict(
        id='GAU-07', chapter='k2', cls='na',
        met='Nicht anwendbar: Zu DEX gehören keine Server. Die Anwendung ist ein Webpart, das im '
            'Browser läuft; es gibt keine Instanz, auf der ein Agent installiert werden könnte.',
        evidence='Architektur, Kapitel 2 „Wo der Code läuft“ — keine Server-Komponente.'),
    dict(
        id='GAU-08', chapter='k2', cls='platform',
        met='Die zugrunde liegende SaaS-Plattform ist Microsoft 365 im Deloitte-Tenant. Deren '
            'Protokollierung ist bereits Teil des Tenants; DEX führt keinen zusätzlichen SaaS-Dienst ein '
            'und legt keine Daten ausserhalb des Tenants ab.',
        evidence='Architektur, Kapitel 1 und 3 — alle Daten liegen in SharePoint Online derselben '
                 'Site Collection.'),
    # ------------------------------------------------------------------ GCM
    dict(
        id='GCM-01', chapter='k15', cls='dex',
        met='Ausführbarer Code besteht aus genau einem Paket: dex-event-platform.sppkg. Es wird über '
            'den App-Katalog des Tenants bereitgestellt — der Katalog ist die Freigabeliste. Die '
            'Fremdbibliotheken im Paket sind in package.json und package-lock.json festgeschrieben und '
            'damit vollständig auflistbar.',
        evidence='Architektur, Kapitel 15 „Deployment und Rollback“; package-lock.json im Repository.',
        open='Jährliche Überprüfung der Bibliotheksliste ist festzulegen; siehe auch GSI-06.'),
    dict(
        id='GCM-04', chapter='k15', cls='dex',
        met='Rückfall ist Teil des Ablaufs: Jede Version wird als eigenes Paket abgelegt und bleibt '
            'abrufbar, die Vorgängerversion ist auf der Downloadseite ausdrücklich als Rollback-Paket '
            'benannt. Ein Rückfall ist das Hochladen des älteren Pakets in den App-Katalog. Die Daten in '
            'den SharePoint-Listen sind davon nicht betroffen — sie werden beim Deployment nicht angefasst.',
        evidence='Architektur, Kapitel 15; Downloadseite der Dokumentation mit Versions- und Rollback-Paket.',
        open='Gegenläufiger Fall: Führt eine Version neue Spalten oder Flags ein, ist der Rückfall '
             'zwar möglich, die neuen Felder bleiben aber in den Listen stehen. Das ist verträglich, '
             'weil die App unbekannte Felder ignoriert — bitte als bewusste Auslegung zur Kenntnis.'),
    dict(
        id='GCM-06', chapter='k17', cls='dex',
        met='Im Quelltext stehen keine Zugangsdaten. Es gibt keine Passwörter, keine API-Schlüssel und '
            'keine Zugriffstoken im Code — die Anwendung authentifiziert ausschliesslich über die '
            'bestehende Sitzung des Nutzers. Der Auslieferungsbuild läuft im Ship-Modus, also ohne '
            'Debug-Artefakte. Testdaten entstehen nur im Demo-Modus und werden zur Laufzeit synthetisch '
            'erzeugt; es werden keine Produktionsdaten kopiert.',
        evidence='Architektur, Kapitel 17 „Sicherheitsbetrachtung“; Prüfung des Quelltextes auf '
                 'Zugangsdaten am 2026-08-12 ohne Befund.'),
    dict(
        id='GCM-07', chapter='k15', cls='flagged',
        met='Es gibt zwei benannte Umgebungen für die Automatisierung, DEV und PROD. Die Kategorisierung beschreibt den Betrieb aber nicht: Die sieben Flows liegen vollständig in der DEV-Umgebung und laufen von dort produktiv gegen die Produktiv-Site. Sie sind fertig und werden nicht mehr weiterentwickelt — die DEV-Umgebung ist damit faktisch die Produktionsumgebung. Für die SPFx-Anwendung existiert gar keine zweite Umgebung; sie liegt allein auf der Produktiv-Site.',
        evidence='Architektur, Kapitel 15 „Deployment und Rollback“ (Ist-Stand und Zielmodell); Auskunft des Betriebs vom 2026-08-17. Nachweis: Power Platform zeigt zwei Umgebungen — GER-DTech-CDE-PowerTeam (produktiv) und GER-DTech-CDE-PowerTeam-DEV (Entwicklung).',
        open='Abweichung, offen benannt: Kategorien sind vergeben, die produktive Last liegt aber in der als Entwicklung kategorisierten Umgebung. Damit ist die Anforderung nicht erfüllt. Zielmodell (vom Betrieb benannt): Flows in die PROD-Umgebung überführen, dort unter einem Service Principal betreiben, der DEV-Umgebung den Zugriff auf Produktivdaten entziehen und eine DEV-Umgebung für die SPFx-Anwendung aufbauen. Damit lösen sich GCM-07 bis GCM-11 und GCM-16 gemeinsam, ebenso GAC-01, GAC-02 und GAC-15. Bitte Zieltermin festhalten.'),
    dict(
        id='GCM-08', chapter='k15', cls='flagged',
        met='Die beiden Power-Platform-Umgebungen sind logisch getrennt und haben eigene Verbindungen und Zugriffsrechte.',
        evidence='Architektur, Kapitel 15 „Deployment und Rollback“ (Ist-Stand und Zielmodell); Auskunft des Betriebs vom 2026-08-17. Nachweis: Power Platform zeigt zwei Umgebungen — GER-DTech-CDE-PowerTeam (produktiv) und GER-DTech-CDE-PowerTeam-DEV (Entwicklung).',
        open='Abweichung: Die Trennung ist nominell, weil die Verbindungen der DEV-Umgebung auf die Produktiv-SharePoint-Site zeigen. Eine Trennung, die auf dieselben Daten greift, trennt nichts, was diese Anforderung schützen soll. Zielmodell (vom Betrieb benannt): Flows in die PROD-Umgebung überführen, dort unter einem Service Principal betreiben, der DEV-Umgebung den Zugriff auf Produktivdaten entziehen und eine DEV-Umgebung für die SPFx-Anwendung aufbauen. Damit lösen sich GCM-07 bis GCM-11 und GCM-16 gemeinsam, ebenso GAC-01, GAC-02 und GAC-15. Bitte Zieltermin festhalten.'),
    dict(
        id='GCM-09', chapter='k15', cls='flagged',
        met='Die DEV-Umgebung hat keine Anwendungsoberfläche — Endnutzer melden sich dort nicht an, es gibt nichts zu bedienen. Zugriff hat der Entwicklungskreis.',
        evidence='Architektur, Kapitel 15 „Deployment und Rollback“ (Ist-Stand und Zielmodell); Auskunft des Betriebs vom 2026-08-17. Nachweis: Power Platform zeigt zwei Umgebungen — GER-DTech-CDE-PowerTeam (produktiv) und GER-DTech-CDE-PowerTeam-DEV (Entwicklung).',
        open='Abweichung im Sinn der Anforderung: Endnutzer greifen zwar nicht selbst zu, werden aber aus dieser Umgebung bedient — jede Bestätigungsmail und jeder Kalendereintrag entsteht dort. Der Schutzzweck (keine produktive Nutzung aus einer Nicht-Produktivumgebung) ist damit nicht erfüllt. Zielmodell (vom Betrieb benannt): Flows in die PROD-Umgebung überführen, dort unter einem Service Principal betreiben, der DEV-Umgebung den Zugriff auf Produktivdaten entziehen und eine DEV-Umgebung für die SPFx-Anwendung aufbauen. Damit lösen sich GCM-07 bis GCM-11 und GCM-16 gemeinsam, ebenso GAC-01, GAC-02 und GAC-15. Bitte Zieltermin festhalten.'),
    dict(
        id='GCM-10', chapter='k15', cls='flagged',
        met='Auf Anwendungsseite gilt: Der Demo-Modus liest keine Produktionsdaten, sondern erzeugt seine Datensätze zur Laufzeit.',
        evidence='Architektur, Kapitel 15 „Deployment und Rollback“ (Ist-Stand und Zielmodell); Auskunft des Betriebs vom 2026-08-17. Nachweis: Power Platform zeigt zwei Umgebungen — GER-DTech-CDE-PowerTeam (produktiv) und GER-DTech-CDE-PowerTeam-DEV (Entwicklung).',
        open='Abweichung, klar benannt: Die Verbindungen der als Nicht-Produktion kategorisierten DEV-Umgebung greifen auf die Produktiv-Site zu und damit auf personenbezogene Teilnehmerdaten. Die Anforderung ist nicht erfüllt. Der Entzug dieses Zugriffs ist ausdrücklich Teil des Zielmodells. Zielmodell (vom Betrieb benannt): Flows in die PROD-Umgebung überführen, dort unter einem Service Principal betreiben, der DEV-Umgebung den Zugriff auf Produktivdaten entziehen und eine DEV-Umgebung für die SPFx-Anwendung aufbauen. Damit lösen sich GCM-07 bis GCM-11 und GCM-16 gemeinsam, ebenso GAC-01, GAC-02 und GAC-15. Bitte Zieltermin festhalten.'),
    dict(
        id='GCM-11', chapter='k15', cls='flagged',
        met='Für die Anwendung ist der Weg in die Produktion festgelegt und dokumentiert: Version setzen, Release Notes, Typprüfung, sauberer Build, Paket ablegen, Bereitstellung über den App-Katalog — nachvollziehbar im Versionsverlauf.',
        evidence='Architektur, Kapitel 15 „Deployment und Rollback“ (Ist-Stand und Zielmodell); Auskunft des Betriebs vom 2026-08-17. Nachweis: Power Platform zeigt zwei Umgebungen — GER-DTech-CDE-PowerTeam (produktiv) und GER-DTech-CDE-PowerTeam-DEV (Entwicklung).',
        open='Für die Flows gibt es heute keinen solchen Weg, weil sie die DEV-Umgebung nie verlassen haben. Die Anforderung greift erst mit dem Zielmodell; dann ist festzulegen, ob die Überführung als Solution-Export und -Import erfolgt. Nur dieser Weg erlaubt einen nachvollziehbaren Rückfall (siehe GCM-04). Zielmodell (vom Betrieb benannt): Flows in die PROD-Umgebung überführen, dort unter einem Service Principal betreiben, der DEV-Umgebung den Zugriff auf Produktivdaten entziehen und eine DEV-Umgebung für die SPFx-Anwendung aufbauen. Damit lösen sich GCM-07 bis GCM-11 und GCM-16 gemeinsam, ebenso GAC-01, GAC-02 und GAC-15. Bitte Zieltermin festhalten.'),
    dict(
        id='GCM-16', chapter='k15', cls='flagged',
        met='Für die Anwendung erfüllt: Produktionsdaten werden nicht für Tests verwendet, der '
            'Demo-Modus erzeugt seine Teilnehmerdaten synthetisch im Browser, und es gibt keinen '
            'Export von Teilnehmerlisten in eine Testumgebung.',
        evidence='Architektur, Kapitel 15 und 17; Funktion buildDemoRegistrations im EventContext; '
                 'Auskunft des Betriebs vom 2026-08-17.',
        open='Abweichung auf der Flow-Seite, die die Anwendungs-Seite nicht aufwiegt: Die Flows liegen '
             'in der DEV-Umgebung und verarbeiten von dort personenbezogene Produktivdaten — Namen und '
             'E-Mail-Adressen der Teilnehmer stehen in jeder Bestätigungsmail, die dort erzeugt wird. '
             'Damit werden Produktivdaten in einer als Nicht-Produktion kategorisierten Umgebung '
             'verarbeitet. Nicht erfüllt, bis das Zielmodell greift (siehe GCM-07 bis GCM-11).'),
    dict(
        id='GCM-18', chapter='k1', cls='na',
        met='Nicht anwendbar im Sinne eigener Netzarchitektur: DEX hat kein Netz, keine Datenbank und '
            'keine Container. Die relevante Architektur ist die Anwendungsarchitektur, und die ist '
            'dokumentiert und bebildert.',
        evidence='Architektur, Kapitel 1 „Das Systembild“ — Bausteine, Datenflüsse und Grenzen.',
        open='Jährliche Überprüfung des Architekturdokuments ist festzulegen.'),
    dict(
        id='GCM-22', chapter='k2', cls='na',
        met='Nicht anwendbar: keine Server, keine Images, kein Betriebssystem im Umfang von DEX.',
        evidence='Architektur, Kapitel 2.'),
    # ------------------------------------------------------------------ GCP
    dict(
        id='GCP-01', chapter='k16', cls='dex',
        met='Die Sicherung liegt vollständig bei SharePoint Online: zweistufiger Papierkorb, '
            'Versionsverlauf auf allen Listen und die Wiederherstellungszusagen von Microsoft 365. '
            'Auf Anwendungsseite kommt hinzu, dass Löschvorgänge fachlich weich sind — eine Abmeldung '
            'setzt den Status, sie löscht die Zeile nicht — und dass jede ausgelieferte Paketversion '
            'aufbewahrt wird, sodass der Anwendungsstand jederzeit zurückgestellt werden kann.',
        evidence='Architektur, Kapitel 16 „Sicherung und Wiederherstellung“.',
        open='Jährliche Überprüfung und ein dokumentierter Wiederherstellungstest sind festzulegen. '
             'Die Aufbewahrungsfristen des Papierkorbs sind Tenant-Vorgabe, keine Projektentscheidung.'),
    # ------------------------------------------------------------------ GIA
    dict(
        id='GIA-01', chapter='k17', cls='dex',
        met='Meldungen der Anwendung nennen den fachlichen Grund, nicht die Technik: „Alle Plätze sind '
            'belegt“, „Für dich ist aktuell kein Programmpunkt freigegeben“. Systeminterna wie Listen-'
            'Kennungen, Pfade oder Statuscodes erscheinen nicht in der Oberfläche. Technische Details '
            'gehen ausschliesslich in die Browser-Konsole der handelnden Person.',
        evidence='Architektur, Kapitel 17.'),
    dict(
        id='GIA-02', chapter='k2', cls='na',
        met='Nicht anwendbar: DEX hat keinen Anmeldevorgang und kein Eingabefeld für Zugangsdaten. '
            'Die Anmeldung erfolgt vor dem Öffnen der Seite durch Entra ID.',
        evidence='Architektur, Kapitel 2.'),
    dict(
        id='GIA-03', chapter='k2', cls='platform',
        met='Erfüllt durch die Bauart: Ein SPFx-Webpart ist Teil einer SharePoint-Seite. Ohne '
            'angemeldete Entra-ID-Sitzung wird die Seite nicht ausgeliefert und der Code nicht '
            'ausgeführt. Einmalanmeldung ist damit nicht optional, sondern Voraussetzung.',
        evidence='Architektur, Kapitel 2.'),
    dict(
        id='GIA-04', chapter='k2', cls='platform',
        met='Mehrfaktor-Authentifizierung greift über die Richtlinien für bedingten Zugriff des '
            'Tenants, bevor die Seite geladen wird. DEX kann sie weder umgehen noch abschwächen.',
        evidence='Architektur, Kapitel 2.'),
    dict(
        id='GIA-05', chapter='k17', cls='dex',
        met='Es gibt in der gesamten Anwendung kein Feld für Benutzername oder Passwort, folglich auch '
            'nichts vorauszufüllen. Angezeigte Namen und Adressen stammen aus dem Profil der bereits '
            'angemeldeten Person beziehungsweise aus der Personenauswahl.',
        evidence='Architektur, Kapitel 17; Prüfung des Quelltextes ohne Befund.'),
    # ------------------------------------------------------------------ GSC
    dict(
        id='GSC-01', chapter='k2', cls='na',
        met='Nicht anwendbar: DEX hat keine eigenen Geräte, Namen oder Topologien. Nach aussen sichtbar '
            'sind ausschliesslich Adressen von SharePoint Online.',
        evidence='Architektur, Kapitel 2.'),
    dict(
        id='GSC-03', chapter='k17', cls='flagged',
        met='Zertifikate für den Transport stellt Microsoft für SharePoint Online; DEX verwaltet keine '
            'eigenen Zertifikate und keine Verschlüsselungsschlüssel.',
        evidence='Architektur, Kapitel 17.',
        open='Offenlegung: Eine Funktion nutzt ein Geheimnis, das nicht über KCM verwaltet wird — der '
             'Selbst-Check-in. Je Event liegt ein Zufallswert in der Event-Zeile, aus dem ein alle paar '
             'Sekunden wechselnder Code für den Aushang berechnet wird (HMAC-SHA-256). Es ist kein '
             'Schlüssel für Daten, sondern ein geteiltes Geheimnis für einen kurzlebigen Anzeigecode. '
             'Wirkung bei Kompromittierung: Ein Unbefugter könnte sich bei genau diesem Event einchecken. '
             'Er ist je Event verschieden und durch Neuerzeugen wechselbar. '
             'WICHTIG zur Verbreitung: Die Funktion ist nicht opt-in. Seit v20.1 erzeugt der erste Klick '
             'auf die QR-Kachel den Wert und schaltet den Selbst-Check-in am Event scharf — die Kachel '
             'erscheint ab fünf Tagen vor Beginn beziehungsweise sobald QR-Codes versendet wurden. Die '
             'Zahl der Events mit einem solchen Geheimnis entspricht deshalb NICHT der Zahl der Events, '
             'die den Selbst-Check-in nutzen (laut Fachbereich in der Regel keines), sondern der Zahl '
             'der Events, bei denen jemand die Kachel einmal geöffnet hat. Empfehlung: Auto-Aktivierung '
             'entfernen (nur noch bewusstes Einschalten) und den Wert nach dem Event löschen. Danach '
             'ist die Abweichung auf die tatsächlich genutzten Events begrenzt und abzählbar.'),
    dict(
        id='GSC-04', chapter='k17', cls='platform',
        met='Alle Daten liegen in SharePoint Online und sind dort im Ruhezustand verschlüsselt; der '
            'Transport läuft über TLS. DEX legt keine Daten ausserhalb ab — kein lokaler Speicher für '
            'personenbezogene Daten, keine Datei, kein Dienst Dritter. Auch Bilder liegen als Inhalt '
            'in SharePoint-Spalten.',
        evidence='Architektur, Kapitel 3 und 17.'),
    dict(
        id='GSC-05', chapter='k17', cls='dex',
        met='Die eingesetzte Verschlüsselung ist im Architekturdokument beschrieben: TLS im Transport, '
            'Verschlüsselung im Ruhezustand durch SharePoint Online, keine eigene Kryptografie ausser '
            'dem unter GSC-03 offengelegten HMAC für den Selbst-Check-in.',
        evidence='Architektur, Kapitel 17 „Sicherheitsbetrachtung“.'),
    dict(
        id='GSC-06', chapter='k17', cls='flagged',
        met='Schlüsselverwaltung für Transport und Ruhezustand liegt bei Microsoft 365.',
        evidence='Architektur, Kapitel 17.',
        open='Siehe GSC-03: Das Geheimnis für den Selbst-Check-in liegt als Feld in der Event-Liste und '
             'ist damit für Organizer dieses Events lesbar. Zusätzlich relevant: Es entsteht heute '
             'automatisch beim ersten Öffnen der QR-Kachel, nicht durch eine bewusste Aktivierung. '
             'Bewertung erbeten; Empfehlung siehe GSC-03.'),
    dict(
        id='GSC-09', chapter='k17', cls='dex',
        met='Eigene Adressen der Anwendung tragen keine personenbezogenen Daten: Sie enthalten die '
            'Event-Nummer und, beim Selbst-Check-in, einen wechselnden Code. Offenlegung der Vollständigkeit '
            'halber: An zwei Stellen erscheint eine E-Mail-Adresse als Parameter, und zwar in Adressen von '
            'Microsoft selbst — beim Profilbild von SharePoint und beim Teams-Chat-Link, der lokal vom '
            'Client verarbeitet wird. Beides sind die vorgesehenen Wege der Plattform, bleiben innerhalb '
            'des Tenants und laufen über TLS.',
        evidence='Architektur, Kapitel 17; geprüfte Stellen: Profilbild-Endpunkt und Teams-Protokolllink.'),
    # ------------------------------------------------------------------ GSI
    dict(
        id='GSI-01', chapter='k17', cls='platform',
        met='DEX betreibt keine Endpunkte, auf denen Schadcode-Erkennung installiert werden könnte. '
            'Relevant ist der eine Weg, auf dem Dateien ins System gelangen: Anhänge, die Teilnehmer oder '
            'Organizer hochladen. Diese landen unmittelbar in SharePoint Online und werden von den '
            'Schutzmechanismen des Tenants geprüft.',
        evidence='Architektur, Kapitel 17.'),
    dict(
        id='GSI-06', chapter='k17', cls='flagged',
        met='Der Anwendungscode selbst enthält keine Server-Komponente; die Angriffsfläche sind die '
            'mitgelieferten Bibliotheken. Ihre Versionen sind festgeschrieben und vollständig auflistbar.',
        evidence='Architektur, Kapitel 17; package-lock.json im Repository. ' + AUDIT_NOTE,
        open='Befund: Die Abhängigkeitsprüfung meldet kritische und hohe Funde. Ein erheblicher Teil '
             'betrifft Build-Werkzeuge, die nicht ausgeliefert werden — das entbindet nicht von der '
             'Bewertung. Nächster Schritt: Aufnahme in den Vulnerability-Management-Service, Abgleich '
             'welche Funde tatsächlich im Bundle landen, dann Aktualisierung oder begründete Annahme. '
             'Ohne diesen Schritt ist die Anforderung nicht erfüllt.'),
    dict(
        id='GSI-09', chapter='k17', cls='open',
        met='Automatisierte Anwendungstests sind bisher nicht durchgeführt worden.',
        evidence='Architektur, Kapitel 17.',
        open='Offen: Anmeldung beim Automated-Application-Testing-Dienst und Durchführung der Scans. '
             'Das kann nur das Projekt beauftragen; die Anwendung ist ein statisches Bundle auf einer '
             'SharePoint-Seite und damit ein üblicher Prüfgegenstand.'),
    # ------------------------------------------------------------------ GSA
    dict(
        id='GSA-04', chapter='k17', cls='open',
        met='Es gibt keine automatisierten Tests im Projekt; geprüft wird über Typprüfung, statische '
            'Codeanalyse im Build und manuelle Abnahme. Eine Sicherheitsprüfung im Sinne dieser '
            'Anforderung hat nicht stattgefunden.',
        evidence='Architektur, Kapitel 17.',
        open='Offen: Sicherheitstest beauftragen und Befunde beheben. Bekannte Ausgangspunkte für den '
             'Test sind in Kapitel 17 benannt — insbesondere die Wiedergabe von organizer-erstelltem '
             'HTML (siehe GSD-02) und die Abhängigkeiten (siehe GSI-06).'),
    # ------------------------------------------------------------------ GSD
    dict(
        id='GSD-01', chapter='k12', cls='dex',
        met='Sicheres Scheitern ist an mehreren Stellen bewusst gebaut. Ein fehlgeschlagener Lesezugriff '
            'wird als „unbekannt“ behandelt und nicht als „nicht vorhanden“ — ein Auswertungslauf '
            'überspringt betroffene Events, statt aus Leere auf Abwesenheit zu schliessen. Vor '
            'unwiderruflichen Löschungen wird zuerst die prüfbare Nebenbuchhaltung aufgeräumt; scheitert '
            'dabei ein Schritt, unterbleibt die Löschung. Rechte werden nie im Browser entschieden, '
            'sondern von SharePoint durchgesetzt — ein manipulierter Client sieht dadurch nicht mehr. '
            'Fehlermeldungen nennen den fachlichen Grund ohne Systeminterna.',
        evidence='Architektur, Kapitel 12 „Wo die Wahrheit liegt“ und Kapitel 17.'),
    dict(
        id='GSD-02', chapter='k17', cls='flagged',
        met='Die üblichen Fehlerklassen sind adressiert. Abfragen an SharePoint maskieren Anführungszeichen '
            'in Filterausdrücken (an über vierzig Stellen konsequent), React maskiert eingesetzte Werte '
            'standardmässig, Berechtigungen werden serverseitig durchgesetzt, und es gibt weder eigene '
            'Authentifizierung noch Sitzungsverwaltung, die man umgehen könnte.',
        evidence='Architektur, Kapitel 17.',
        open='Offenlegung eines Restrisikos: An fünf Stellen wird gespeichertes HTML unmaskiert angezeigt '
             '— Event-Beschreibungen, Feldbeschreibungen und Mailvorlagen, die Organizer in einem '
             'Rich-Text-Editor erfassen. Vor der Ausgabe werden Skript- und Stilblöcke, Ereignis-Attribute '
             'und javascript:-Verweise entfernt. Das ist eine Ausschlussliste, keine Freigabeliste, und '
             'damit prinzipiell umgehbar. Mildernd: Erfassen können solche Inhalte nur angemeldete, '
             'namentlich bekannte Organizer desselben Tenants; jede Änderung ist protokolliert; die '
             'Wiedergabe erfolgt nur gegenüber Nutzern desselben Tenants. Empfehlung: Umstellung auf einen '
             'Bereiniger mit Freigabeliste. Als Befund im Sicherheitstest (GSA-04) einplanen.'),
    # ------------------------------------------------------------------ GWAF
    dict(
        id='GWAF-01', chapter='k2', cls='na',
        met='Nicht anwendbar: DEX stellt keinen eigenen Webdienst bereit, der vorgelagert geschützt '
            'werden könnte. Ausgeliefert wird eine JavaScript-Datei über SharePoint Online; sämtlicher '
            'Verkehr läuft gegen Microsoft-Endpunkte, die hinter deren eigenem Schutz stehen. Ein WAF im '
            'Verantwortungsbereich des Projekts gäbe es nicht vorzuschalten.',
        evidence='Architektur, Kapitel 1 und 2.'),
]

CLS_LABEL = {
    'platform': ('Plattform', 'Erfüllt durch Microsoft 365 / Deloitte-Tenant'),
    'dex': ('DEX', 'Erfüllt durch die Anwendung'),
    'na': ('Nicht anwendbar', 'Trifft auf diese Bauart nicht zu'),
    'open': ('Offen', 'Bestätigung durch Betrieb oder GTS nötig'),
    'flagged': ('Zur Bewertung', 'Abweichung oder Restrisiko, offengelegt'),
}
