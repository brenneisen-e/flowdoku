# Power Automate Flow Definitionen

## Übersicht

| Flow | Trigger | Zweck |
|------|---------|-------|
| DEX_IDReorder_TeilnehmerIDs | Neuer Eintrag in DEX_IDReorder | TeilnehmerIDs neu vergeben + Nachrücken |
| DEX_Outlook_Einladungen | Neuer Eintrag in DEX_Outlook | Outlook-Kalender Einladen/Ausladen |
| DEX_Emails_Senden | Neuer Eintrag in DEX_Emails | E-Mails versenden |

---

## Flow 1: DEX_IDReorder_TeilnehmerIDs

**Bitte den aktuellen Flow-JSON hier einfügen.**

Letzte bekannte Struktur:

```
Trigger: When item created in DEX_IDReorder
→ Get_Event_Details (DEX_Events by EventId)
→ Init_RealEventId (empty)
→ Init_Attendees (empty)
→ Settings (siteAddress, listName, batchSize)
→ Get_ListItemType (ListItemEntityTypeFullName)
→ Get_Active_Participants (filter: Angemeldet/QR versendet/Eingecheckt)
→ Generate_Indices
→ GenerateSPData (ID + TeilnehmerID)
→ BatchGuids
→ batchTemplate (PATCH mit __metadata hardcoded)
→ Loop_Batches → Select_map → batchData → SendBatch
→ Get_EventDetails (MaxParticipants)
→ Check_Nachrücken (Angemeldet < MaxParticipants?)
  → Yes:
    → Get_Waitlist_First (erster Warteliste-Eintrag)
    → Condition_1 (length > 0?)
      → Yes:
        → Promote_Waitlist (MERGE: Status=Angemeldet)
        → Get_Email_Template (aus DEX_EmailTemplates)
        → Queue Email (in DEX_Emails)
        → Queue Outlook (in DEX_Outlook: Einladen)
→ Set_Done (DEX_IDReorder Status=Done)
→ Error_Handler → Set_Failed
```

### Bekannte Fixes:
- batchTemplate: changeSetGUID aus BatchGuids (nicht guid())
- __metadata hardcoded im Template (nicht über Select)
- Select_map: replace |ID| und |TID| statt |RowData|
- Get_Waitlist_First Response: d.results (nicht value)
- Promote_Waitlist: odata=verbose + __metadata + odata-version: ''
- Condition_1: coalesce mit d.results
- Queue Email: Body aus DEX_EmailTemplates mit Platzhalter-Ersetzung

---

## Flow 2: DEX_Outlook_Einladungen

**Bitte den aktuellen Flow-JSON hier einfügen.**

Letzte bekannte Struktur:

```
Trigger: When item created in DEX_Outlook (Concurrency: 1)
→ Get_Event_Details (DEX_Events by EventId)
→ Init_RealEventId (empty string)
→ Init_Attendees (empty array)
→ Has_OutlookEventId? (length CalendarLink > 0)
  → Yes:
    → Find_Outlook_Event (Graph API: GET events?$filter=iCalUId)
    → Set varRealEventId (first result id)
    → Check_EventFound (length varRealEventId > 0?)
      → Yes:
        → Get_Existing_Event (Graph API: GET event details)
        → Set var_Attendees (attendees array)
        → Check_ActionType (Einladen?)
          → Yes (Einladen):
            → Add_Attendee (append to array)
            → Update_Event_Einladen (Graph API: PATCH attendees)
          → No (Ausladen):
            → Filter_Attendees (remove by email)
            → Update_Event_Ausladen (Graph API: PATCH filtered)
        → Set_Sent (DEX_Outlook Status=Sent)
        → Set_Failed_1_1 (Run After: Failed)
      → No:
        → Set_Failed_1
  → No:
    → Set_Failed
```

### Bekannte Fixes:
- CalendarLink statt OutlookEventId (iCalUId dort gespeichert)
- Graph API via Office 365 Outlook Connector (nicht SharePoint)
- Concurrency: 1 (verhindert Race Conditions)

---

## Flow 3: DEX_Emails_Senden

**Bitte den aktuellen Flow-JSON hier einfügen.**

Erwartete Struktur:

```
Trigger: When item created in DEX_Emails
→ Update Status to Processing
→ Send Email (via Office 365 / Send Mail)
  - To: Recipient
  - Subject: Title
  - Body: Body (HTML)
→ Set Status to Sent
→ Error: Set Status to Failed
```

---

## SharePoint Listen-Referenz

| Liste | Zweck | Wichtige Spalten |
|-------|-------|-----------------|
| DEX_Events | Event-Stammdaten | Title, EventType, Location, StartDate, EndDate, MaxParticipants, Organizer, EmailLanguage, EmailTemplateOverrides, CalendarLink, SubsiteUrl |
| DEX_Roles | Rollenverwaltung | Title (Email), UserName, Role, UserLocation |
| DEX_Emails | Email-Queue | Title (Subject), Recipient, RecipientName, Body, EmailType, EventTitle, Status |
| DEX_Outlook | Outlook-Queue | Title, Attendee, EventId, ActionType (Einladen/Ausladen), Status |
| DEX_IDReorder | ID-Neuvergabe Queue | Title, EventId, EventNumber, SubsiteUrl, Status |
| DEX_Participants | Zentrale Teilnehmerliste | Title (Email), Vorname, Nachname, Email, EventRegistered, EventOnWaitlist |
| DEX_EmailTemplates | Email-Vorlagen (DE+EN) | Title, TemplateType, Language, Subject, HeadingColor, Heading, BodyHtml |
| Teilnehmer (pro Subsite) | Registrierungen pro Event | Title (Email), TeilnehmerID, Vorname, Nachname, ParticipantEmail, Status, RegistrationDate, CancellationDate, CustomData |

---

## Berechtigungen (Visitors = DEALL)

| Liste | Visitors (DEALL) | Owners (Admins) | Organizer |
|-------|-----------------|-----------------|-----------|
| DEX_Events | Read | Full Control | Contribute (individuell) |
| DEX_Roles | - | Full Control | Read |
| DEX_Emails | Contribute + ILS | Full Control | - |
| DEX_Outlook | Contribute + ILS | Full Control | - |
| DEX_IDReorder | Contribute + ILS | Full Control | - |
| DEX_Participants | Contribute + ILS | Full Control | - |
| DEX_EmailTemplates | Read | Full Control | - |
| Event Subsite | Read | Full Control | Full Control |
| Teilnehmer-Liste | Contribute + ILS | Full Control | Full Control |
