# Power Automate Flow Definitionen

## Übersicht

| Flow | Trigger | Zweck |
|------|---------|-------|
| DEX_IDReorder_TeilnehmerIDs | Neuer Eintrag in DEX_IDReorder | TeilnehmerIDs neu vergeben + Nachrücken |
| DEX_Outlook_Einladungen | Neuer Eintrag in DEX_Outlook | Outlook-Kalender Einladen/Ausladen |
| DEX_Emails_Senden | Neuer Eintrag in DEX_Emails | E-Mails versenden |

---

## Flow 1: DEX_IDReorder_TeilnehmerIDs

### Aktueller Stand (2026-04-06)

#### Trigger
```json
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "9d46ff77-5fe2-4e1d-9b93-14b9dca1a360"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "GetOnNewItems"
    }
  },
  "recurrence": { "frequency": "Minute", "interval": 1 },
  "runtimeConfiguration": { "concurrency": { "runs": 1, "maximumWaitingRuns": 100 } },
  "splitOn": "@triggerOutputs()?['body/value']"
}
```

#### Update_item (Status → Processing)
```json
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "9d46ff77-5fe2-4e1d-9b93-14b9dca1a360",
      "id": "@triggerBody()?['ID']",
      "item/Title": "@triggerBody()?['Title']",
      "item/Status/Value": "Processing"
    },
    "host": { "operationId": "PatchItem" }
  },
  "runAfter": {}
}
```

#### Settings
```json
{
  "type": "Compose",
  "inputs": {
    "siteAddress": "@{triggerOutputs()?['body/SubsiteUrl']}",
    "listName": "Teilnehmer",
    "batchSize": 250
  },
  "runAfter": { "Update_item": ["Succeeded"] }
}
```

#### Get_ListItemType
```json
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "@outputs('Settings')?['siteAddress']",
      "parameters/method": "GET",
      "parameters/uri": "_api/web/lists/getbytitle('@{outputs('Settings')?['listName']}')?$select=ListItemEntityTypeFullName"
    },
    "host": { "operationId": "HttpRequest" }
  },
  "runAfter": { "Settings": ["Succeeded"] }
}
```

#### Get_Active_Participants
```json
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "@outputs('Settings')?['siteAddress']",
      "parameters/method": "GET",
      "parameters/uri": "_api/web/lists/getbytitle('@{outputs('Settings')?['listName']}')/items?$select=Id,TeilnehmerID,Status,RegistrationDate&$filter=(Status eq 'Angemeldet') or (Status eq 'QR versendet') or (Status eq 'Eingecheckt')&$orderby=RegistrationDate asc&$top=5000",
      "parameters/headers": { "Accept": "application/json;odata=nometadata" }
    },
    "host": { "operationId": "HttpRequest" }
  },
  "runAfter": { "Get_ListItemType": ["Succeeded"] }
}
```

#### Generate_Indices
```json
{
  "type": "Compose",
  "inputs": "@range(0, length(body('Get_Active_Participants')?['value']))",
  "runAfter": { "Get_Active_Participants": ["Succeeded"] }
}
```

#### GenerateSPData
```json
{
  "type": "Select",
  "inputs": {
    "from": "@outputs('Generate_Indices')",
    "select": {
      "ID": "@body('Get_Active_Participants')?['value'][item()]?['Id']",
      "TeilnehmerID": "@add(item(), 1)"
    }
  },
  "runAfter": { "Generate_Indices": ["Succeeded"] }
}
```

#### BatchGuids
```json
{
  "type": "Compose",
  "inputs": { "batchGUID": "@{guid()}", "changeSetGUID": "@{guid()}" },
  "runAfter": { "GenerateSPData": ["Succeeded"] }
}
```

#### batchTemplate
```json
{
  "type": "Compose",
  "inputs": "--changeset_@{outputs('BatchGuids')?['changeSetGUID']}\nContent-Type: application/http\nContent-Transfer-Encoding: binary\n\nPATCH @{outputs('Settings')?['siteAddress']}/_api/web/lists/getByTitle('@{outputs('Settings')?['listName']}')/items(|ID|) HTTP/1.1\nContent-Type: application/json;odata=verbose\nAccept: application/json;odata=verbose\nIf-Match: *\n\n{\"__metadata\":{\"type\":\"@{body('Get_ListItemType')?['d']?['ListItemEntityTypeFullName']}\"},\"TeilnehmerID\":|TID|}\n",
  "runAfter": { "BatchGuids": ["Succeeded"] }
}
```

#### Process_Batch Scope
```json
{
  "type": "Scope",
  "actions": {
    "Loop_Batches": {
      "type": "Foreach",
      "foreach": "@chunk(body('GenerateSPData'), outputs('Settings')?['batchSize'])",
      "actions": {
        "Select_map": {
          "type": "Select",
          "inputs": {
            "from": "@items('Loop_Batches')",
            "select": "@replace(replace(outputs('batchTemplate'), '|ID|', string(item()?['ID'])), '|TID|', string(item()?['TeilnehmerID']))"
          }
        },
        "batchData": {
          "type": "Compose",
          "inputs": "@join(body('Select_map'), decodeUriComponent('%0A'))",
          "runAfter": { "Select_map": ["Succeeded"] }
        },
        "SendBatch": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "dataset": "@outputs('Settings')?['siteAddress']",
              "parameters/method": "POST",
              "parameters/uri": "_api/$batch",
              "parameters/headers": {
                "Content-Type": "multipart/mixed;boundary=batch_@{outputs('BatchGuids')?['batchGUID']}"
              },
              "parameters/body": "--batch_@{outputs('BatchGuids')?['batchGUID']}\nContent-Type: multipart/mixed; boundary=\"changeset_@{outputs('BatchGuids')?['changeSetGUID']}\"\nContent-Length: @{length(outputs('batchData'))}\nContent-Transfer-Encoding: binary\n\n@{outputs('batchData')}\n--changeset_@{outputs('BatchGuids')?['changeSetGUID']}--\n\n--batch_@{outputs('BatchGuids')?['batchGUID']}--\n"
            },
            "host": { "operationId": "HttpRequest" }
          },
          "runAfter": { "batchData": ["Succeeded"] }
        }
      },
      "runtimeConfiguration": { "concurrency": { "repetitions": 1 } }
    },
    "Get_EventDetails": {
      "type": "OpenApiConnection",
      "inputs": {
        "parameters": {
          "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
          "table": "28457815-1163-4e92-8b08-3ae43f477d9e",
          "$filter": "ID eq '@{triggerOutputs()?['body/EventId']}'",
          "$top": 1
        },
        "host": { "operationId": "GetItems" }
      },
      "runAfter": { "Loop_Batches": ["Succeeded"] }
    },
    "Check_Nachrücken": {
      "type": "If",
      "expression": {
        "and": [
          { "less": ["@length(body('GenerateSPData'))", "@first(outputs('Get_EventDetails')?['body/value'])?['MaxParticipants']"] },
          { "greater": ["@first(outputs('Get_EventDetails')?['body/value'])?['MaxParticipants']", 0] }
        ]
      },
      "actions": {
        "Get_Waitlist_First": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "dataset": "@outputs('Settings')?['siteAddress']",
              "parameters/method": "GET",
              "parameters/uri": "_api/web/lists/getbytitle('@{outputs('Settings')?['listName']}')/items?$select=Id,Vorname,Nachname,ParticipantName,ParticipantEmail&$filter=Status eq 'Warteliste'&$orderby=RegistrationDate asc&$top=1"
            },
            "host": { "operationId": "HttpRequest" }
          }
        },
        "Condition_1": {
          "type": "If",
          "expression": { "and": [{ "greater": ["@length(coalesce(body('Get_Waitlist_First')?['d']?['results'], json('[]')))", 0] }] },
          "actions": {
            "Promote_Waitlist": {
              "type": "OpenApiConnection",
              "inputs": {
                "parameters": {
                  "dataset": "@outputs('Settings')?['siteAddress']",
                  "parameters/method": "POST",
                  "parameters/uri": "_api/web/lists/getbytitle('@{outputs('Settings')?['listName']}')/items(@{first(body('Get_Waitlist_First')?['d']?['results'])?['Id']})",
                  "parameters/headers": {
                    "Content-Type": "application/json;odata=verbose",
                    "IF-MATCH": "*",
                    "X-HTTP-Method": "MERGE",
                    "Accept": "application/json;odata=verbose"
                  },
                  "parameters/body": "{\"__metadata\":{\"type\":\"SP.Data.TeilnehmerListItem\"},\"Status\":\"Angemeldet\",\"TeilnehmerID\":@{add(length(body('GenerateSPData')), 1)}}"
                },
                "host": { "operationId": "HttpRequest" }
              }
            },
            "Get_Email_Template": {
              "type": "OpenApiConnection",
              "inputs": {
                "parameters": {
                  "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                  "parameters/method": "GET",
                  "parameters/uri": "@concat('_api/web/lists/getbytitle(''DEX_EmailTemplates'')/items?$filter=TemplateType eq ''Nachruecken'' and Language eq ''', coalesce(first(outputs('Get_EventDetails')?['body/value'])?['EmailLanguage'], 'EN'), '''&$select=Subject,BodyHtml&$top=1')"
                },
                "host": { "operationId": "HttpRequest" }
              },
              "runAfter": { "Promote_Waitlist": ["Succeeded"] }
            },
            "Queue_Email": {
              "type": "OpenApiConnection",
              "inputs": {
                "parameters": {
                  "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                  "table": "57aa0840-df98-41ae-a39b-323c0b80ae3b",
                  "item/Title": "@replace(replace(coalesce(first(body('Get_Email_Template')?['d']?['results'])?['Subject'], concat('Platz frei: ', triggerOutputs()?['body/Title'])), '{{Name}}', first(body('Get_Waitlist_First')?['d']?['results'])?['ParticipantName']), '{{EventTitle}}', triggerOutputs()?['body/Title'])",
                  "item/Recipient": "@first(body('Get_Waitlist_First')?['d']?['results'])?['ParticipantEmail']",
                  "item/RecipientName": "@first(body('Get_Waitlist_First')?['d']?['results'])?['ParticipantName']",
                  "item/EmailType/Value": "Nachruecken",
                  "item/EventTitle": "@triggerOutputs()?['body/Title']",
                  "item/Status/Value": "Pending",
                  "item/Body": "@replace(replace(coalesce(first(body('Get_Email_Template')?['d']?['results'])?['BodyHtml'], concat('Hallo ', first(body('Get_Waitlist_First')?['d']?['results'])?['ParticipantName'], ', du bist nachgerückt!')), '{{Name}}', first(body('Get_Waitlist_First')?['d']?['results'])?['ParticipantName']), '{{EventTitle}}', triggerOutputs()?['body/Title'])",
                  "item/EventId": "@triggerOutputs()?['body/EventId']"
                },
                "host": { "operationId": "PostItem" }
              },
              "runAfter": { "Get_Email_Template": ["SUCCEEDED"] }
            },
            "Queue_Outlook": {
              "type": "OpenApiConnection",
              "inputs": {
                "parameters": {
                  "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                  "table": "d794655b-c950-416c-a478-5dbae285e46d",
                  "item/Title": "Einladen: @{triggerOutputs()?['body/Title']}",
                  "item/Attendee": "@first(body('Get_Waitlist_First')?['d']?['results'])?['ParticipantEmail']",
                  "item/EventId": "@triggerOutputs()?['body/EventId']",
                  "item/ActionType/Value": "Einladen",
                  "item/Status/Value": "Pending"
                },
                "host": { "operationId": "PostItem" }
              },
              "runAfter": { "Queue_Email": ["Succeeded"] }
            }
          },
          "else": { "actions": {} },
          "runAfter": { "Get_Waitlist_First": ["Succeeded"] }
        }
      },
      "else": { "actions": {} },
      "runAfter": { "Get_EventDetails": ["Succeeded"] }
    },
    "DEX_IDReorder": {
      "type": "OpenApiConnection",
      "inputs": {
        "parameters": {
          "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
          "table": "9d46ff77-5fe2-4e1d-9b93-14b9dca1a360",
          "id": "@triggerBody()?['ID']",
          "item/Title": "@triggerBody()?['Title']",
          "item/Status/Value": "Done"
        },
        "host": { "operationId": "PatchItem" }
      },
      "runAfter": { "Check_Nachrücken": ["Succeeded"] }
    },
    "Error_Handler": {
      "type": "Scope",
      "actions": {
        "Set_Failed": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
              "table": "9d46ff77-5fe2-4e1d-9b93-14b9dca1a360",
              "id": "@triggerBody()?['ID']",
              "item/Title": "@triggerBody()?['Title']",
              "item/Status/Value": "Failed"
            },
            "host": { "operationId": "PatchItem" }
          }
        }
      },
      "runAfter": { "DEX_IDReorder": ["Failed"] }
    }
  },
  "runAfter": { "batchTemplate": ["Succeeded"] }
}
```

### Listen-GUIDs (aktuell)
| Liste | GUID |
|-------|------|
| DEX_IDReorder | 9d46ff77-5fe2-4e1d-9b93-14b9dca1a360 |
| DEX_Events | 28457815-1163-4e92-8b08-3ae43f477d9e |
| DEX_Emails | 57aa0840-df98-41ae-a39b-323c0b80ae3b |
| DEX_Outlook | d794655b-c950-416c-a478-5dbae285e46d |

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

### Offene Änderungen:
- CalendarLink statt OutlookEventId verwenden (iCalUId dort gespeichert)
- Graph API via Office 365 Outlook Connector (nicht SharePoint)
- Concurrency: 1 (verhindert Race Conditions)

---

## Flow 3: DEX_Emails_Senden

**Bitte den aktuellen Flow-JSON hier einfügen.**

Erwartete Struktur:

```
Trigger: When item created in DEX_Emails
→ Update Status to Processing
→ Send Email from shared mailbox (no_reply.events@deloitte.de)
  - To: Recipient
  - Subject: Title
  - Body: Body (HTML)
→ Set Status to Sent
→ Error: Set Status to Failed
```

### Wichtig:
- Shared Mailbox: no_reply.events@deloitte.de
- Body ist bereits fertig formatiertes HTML (aus App oder Template)
- Concurrency: 1 empfohlen

---

## SharePoint Listen-Referenz

| Liste | Zweck | Wichtige Spalten |
|-------|-------|-----------------|
| DEX_Events | Event-Stammdaten | Title, EventType, Location, StartDate, EndDate, MaxParticipants, Organizer, EmailLanguage, EmailTemplateOverrides, CalendarLink, SubsiteUrl |
| DEX_Roles | Rollenverwaltung | Title (Email), UserName, Role, UserLocation |
| DEX_Emails | Email-Queue | Title (Subject), Recipient, RecipientName, Body, EmailType, EventTitle, EventId, Status |
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
