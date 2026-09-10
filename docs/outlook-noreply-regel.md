# Outlook-Regel: Autoantwort für `no_reply.events@deloitte.de`

Stand 10.09.2026. Anlass: Das Versand-Postfach der DEX-Mails bekommt Antworten
von Teilnehmern (Inbox 6.513 ungelesen). Wer dort hinschreibt, bekommt heute
nichts zurück — die Frage versandet, und die Person glaubt, sie habe sich
gemeldet.

Ziel: Wer **selbst** an das Postfach schreibt, bekommt einmalig eine Antwort
mit dem Hinweis, dass hier niemand mitliest, plus dem Weg über DEX. Was der
Kalender und andere Automaten dorthin schicken, bleibt still.

Nicht antworten auf:
- Terminzusagen (**Accepted**), Absagen (**Declined**), Vorbehalte (**Tentative**)
- Abwesenheitsnotizen (**Automatic reply** / **Automatische Antwort**)
- Termin-Weiterleitungen (**Meeting Forward Notification**), Absagen von
  Terminen, Unzustellbarkeitsberichte

---

## Die Regel auf einen Blick

| # | Teil | Einstellung |
|---|------|-------------|
| 1 | Gilt für | Alle eingehenden Nachrichten im Postfach `no_reply.events@deloitte.de` |
| 2 | Aktion | **have server reply using a specific message** (serverseitig, antwortet je Absender nur EINMAL) |
| 3 | Ausnahme A | **except if it uses the _form name_ form** → die Kalender-Formulare |
| 4 | Ausnahme B | **except if it is an automatic reply** |
| 5 | Ausnahme C | **except if the subject contains specific words** → Sicherheitsnetz für den Rest |

---

## Vorher: Die Regel muss dem SHARED MAILBOX gehören

Das ist die Falle, an der es sonst scheitert. Legst du die Regel an, während
dein eigenes Postfach aktiv ist, landet sie in **deinem** Postfach und feuert
nie für `no_reply.events`.

- [ ] Outlook öffnen, **File** → **Manage Rules & Alerts**.
- [ ] Oben im Dialog steht **Apply changes to this folder:** mit einer
      Auswahl. Dort **no_reply.events@deloitte.de** wählen.
- [ ] Steht das Postfach dort **nicht** zur Auswahl (typisch bei
      automatisch eingebundenen Shared Mailboxes), musst du es einmal als
      eigenes Konto einbinden: **File** → **Account Settings** →
      **Account Settings** → **New** → die Adresse eintragen. Oder du machst
      es über die Weboberfläche (siehe „Wenn es im Desktop nicht geht").

> Warum das wichtig ist: Nur eine Regel, die IM Postfach liegt und
> **serverseitig** läuft, wirkt auch dann, wenn niemand Outlook offen hat.
> Genau das ist bei einem Versand-Postfach der Normalfall.

---

## Schritt für Schritt

### 1 — Regel anlegen

- [ ] Im Dialog **Rules and Alerts** auf **New Rule…**
- [ ] Unter **Start from a blank rule** die Zeile
      **Apply rule on messages I receive** wählen → **Next**
- [ ] **Keine** Bedingung ankreuzen → **Next**
- [ ] Outlook fragt nach: *„This rule will be applied to every message you
      receive. Is this correct?"* → **Yes**

### 2 — Die Antwort festlegen

- [ ] Bei den Aktionen **have server reply using a specific message**
      ankreuzen
- [ ] Unten auf den unterstrichenen Text **a specific message** klicken —
      es öffnet sich ein leeres Nachrichtenfenster
- [ ] **Subject** eintragen:

```
Dieses Postfach wird nicht gelesen - so erreichst du uns
```

- [ ] Als Text den Block aus dem Abschnitt „Der Antworttext" unten einfügen
- [ ] Das Fenster über **Save & Close** schließen (NICHT senden — es gibt
      keinen Empfänger, das ist richtig so)
- [ ] **Next**

> **Nicht** die Aktion **reply using a specific template** nehmen. Die läuft
> nur, solange Outlook auf einem Rechner offen ist. „Server reply" läuft im
> Exchange und antwortet jedem Absender **einmal** — kein Ping-Pong mit
> Abwesenheitsnotizen.

### 3 — Ausnahme A: Kalender-Nachrichten

- [ ] **except if it uses the _form name_ form** ankreuzen
- [ ] Unten auf **form name** klicken
- [ ] Im Dialog **Choose Forms** oben in der Auswahl **Application Forms**
      wählen
- [ ] Diese Einträge nacheinander markieren und mit **Add** übernehmen:
      - **Accept Meeting Response**
      - **Decline Meeting Response**
      - **Tentative Meeting Response**
      - **Meeting Cancellation**
      - **Meeting Request**
      - **Meeting Update**
      - **Meeting Forward Notification** *(falls vorhanden — sonst fängt
        Ausnahme C sie ab)*
- [ ] **Close**

> Warum über das Formular und nicht über den Betreff: Der Betreff ist
> übersetzt. Dein Screenshot zeigt beides nebeneinander — „Accepted: DTP
> Basics Training" und „Abgelehnt: P/D Meeting T&T+". Das Formular ist bei
> beiden dasselbe.
>
> Und warum nicht über den Nachrichtenkopf (`Content-Class`): Hausinterne
> Nachrichten haben oft gar keine Internet-Header. Eine Kopf-Regel würde
> ausgerechnet bei Kolleginnen und Kollegen nicht greifen.

### 4 — Ausnahme B: Abwesenheitsnotizen

- [ ] **except if it is an automatic reply** ankreuzen

### 5 — Ausnahme C: Sicherheitsnetz über den Betreff

- [ ] **except if the subject contains specific words** ankreuzen
- [ ] Auf **specific words** klicken und diese Begriffe **einzeln** eintragen
      (nach jedem **Add**):

```
Accepted:
Declined:
Tentative:
Canceled:
Zugesagt:
Abgelehnt:
Mit Vorbehalt:
Abgesagt:
Automatic reply:
Automatische Antwort:
Meeting Forward Notification
Undeliverable:
Unzustellbar:
```

- [ ] **OK** → **Next**

> Diese Liste ist bewusst das ZWEITE Netz, nicht das erste: Sie hängt an
> Übersetzungen und geht kaputt, sobald jemand einen anderen Client benutzt.
> Ausnahme A ist die belastbare.

### 6 — Fertigstellen

- [ ] Namen vergeben, z.B. `DEX no_reply - Hinweis an Absender`
- [ ] **Turn on this rule** muss angehakt sein
- [ ] **Run this rule now on messages already in "Inbox"** NICHT anhaken —
      sonst gehen 6.513 Antworten raus
- [ ] **Finish**, dann **OK**

---

## Der Antworttext

Kopierfertig. Duzen, wie überall in DEX.

```
Hallo,

danke für deine Nachricht — und Entschuldigung für die unpersönliche Antwort.

Dieses Postfach verschickt nur die Mails der DEX Event Experience Platform.
Es wird nicht gelesen, und deine Nachricht wird auch nicht weitergeleitet.

So kommst du weiter:

1. Anmelden, abmelden, Angaben oder Termine ändern kannst du selbst in DEX:
   https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView

2. Du hast eine Frage zu einem bestimmten Event? Öffne das Event in DEX —
   auf der Event-Seite stehen die Organizer mit Namen. Schreib ihnen direkt,
   sie können dir am schnellsten helfen.

3. Du weißt nicht mehr, um welches Event es ging? Unter "Meine Events" in DEX
   findest du alles, wofür du angemeldet bist — mit den passenden Organizern.

Viele Grüße
Dein DEX-Team
```

---

## Testen

- [ ] Von deinem eigenen Postfach eine ganz normale Mail an
      `no_reply.events@deloitte.de` schicken → die Antwort muss innerhalb
      weniger Minuten zurückkommen.
- [ ] Ein **zweites** Mal von derselben Adresse schreiben → es kommt
      **keine** zweite Antwort. Das ist richtig und beabsichtigt: „Server
      reply" antwortet jedem Absender nur einmal.
- [ ] Eine Termineinladung an dich selbst legen, bei der
      `no_reply.events@deloitte.de` eingeladen ist, und zusagen → es darf
      **keine** Antwort kommen.

| Beobachtung | Ursache |
|---|---|
| Gar keine Antwort, auch nicht beim ersten Versuch | Die Regel liegt in deinem eigenen Postfach statt im Shared Mailbox — Schritt „Apply changes to this folder" prüfen |
| Antwort kommt nur, wenn Outlook offen ist | Es wurde **reply using a specific template** statt **have server reply using a specific message** gewählt |
| Zusagen lösen doch eine Antwort aus | In **Choose Forms** stand die Auswahl nicht auf **Application Forms**, oder ein Formular fehlt in der Liste |
| Zweite Mail desselben Absenders bekommt keine Antwort | Kein Fehler — siehe oben. Exchange merkt sich den Absender, bis die Regel geändert oder neu eingeschaltet wird |

---

## Wenn es im Desktop nicht geht

Die Weboberfläche (Outlook im Browser, Shared Mailbox über **Open another
mailbox**) kann Regeln — aber **keine** Autoantwort als Regel-Aktion. Dort
gibt es nur „Automatische Antworten" für das ganze Postfach, und die
unterscheidet nicht zwischen einer echten Frage und einer Terminzusage. Für
diesen Zweck also ungeeignet.

## Die eigentlich bessere Lösung (braucht Exchange-Admin)

Eine **Mail flow rule** (Transportregel) im Exchange Admin Center ist dem
Regel-Bastelwerk oben überlegen:

- Sie läuft im Transport, unabhängig von Postfach und Client.
- Sie kennt eine saubere Bedingung **The message type is** mit den Werten
  **Calendaring** und **Automatic reply** — damit entfallen die Ausnahmen A
  und C komplett; kein Übersetzungs-Ratespiel mehr.
- Aktion **Notify the sender with a message** stellt zu und schickt den
  Hinweis, Aktion **Reject the message with the explanation** blockt
  stattdessen.

Regel-Skizze für die IT:
`If recipient is no_reply.events@deloitte.de` **and** `message type is not
Calendaring` **and** `message type is not Automatic reply` → `Notify the
sender with the explanation …`

## Und die Ursache statt des Symptoms

In der DEX-Mailvorlage (`services/EmailTemplates.ts`, Fußzeile) steht heute
nur „Made with DEX App" und die Rechtszeile — **kein** Hinweis, dass das
Postfach nicht gelesen wird. Eine Zeile dort erreicht jeden Empfänger, bevor
er antwortet; die Regel erreicht ihn erst danach. Beides zusammen ist die
vollständige Lösung, die Fußzeile ist der wirksamere Teil.
