# Power Automate Flow JSONs

Dieses Dokument enthält die vollständigen Flow-Definitionen aller **6 DEX-Flows**.
Wird aktualisiert wenn Flows geändert werden.

Übersicht:
1. **DEX_IDReorder_TeilnehmerIDs** — TeilnehmerIDs renummerieren + Warteliste nachrücken
2. **DEX_SEND_MAIL** — Mail-Versand aus DEX_Emails-Queue
3. **DEX_CreateOutlookEvent** — Outlook-Termin initial anlegen
4. **DEX_Outlook_Einladungen** — Teilnehmer zum Outlook-Termin hinzufügen/entfernen, Termin aktualisieren/löschen
5. **DEX_OutlookDeclineHandler** — Decline-Mails abfangen → Reminder-Mail queuen (inkl. weitergeleitete Declines FW:/WG:)
6. **DEX_OutlookForwardHandler** (NEU) — Meeting-Forward-Notifications abfangen → FYI-Mail an Organizer wenn weitergeleiteter Empfänger nicht registriert

---

## 1. DEX_IDReorder_TeilnehmerIDs

**Trigger:** Neuer Eintrag in DEX_IDReorder
**Zweck:** TeilnehmerIDs neu vergeben (Aktive UND Warteliste) + Nachrücken von Warteliste
**Letztes Update:** 2026-04-22 (v6.6)

**Änderungen 2026-04-22 (v6.6) — Zwei-Pass-Sortierung:**

Neuer Compose-Step `Sort_ByStatusPriority` zwischen `Count_Active` und `Generate_Indices`.
Er liefert die Enrolled-Items in der gewünschten Reihenfolge: erst alle Angemeldeten
(Status ∈ Angemeldet / QR versendet / Eingecheckt, sortiert nach RegistrationDate), dann
alle Warteliste-Teilnehmer (auch nach RegistrationDate). Die Expression ist:

```
@union(
  filter(body('Get_Enrolled_Participants')?['value'], not(equals(item()?['Status'], 'Warteliste'))),
  filter(body('Get_Enrolled_Participants')?['value'], equals(item()?['Status'], 'Warteliste'))
)
```

`Generate_Indices` zählt die Länge von `Sort_ByStatusPriority` statt von
`Get_Enrolled_Participants`. `GenerateSPData` greift bei `ID` auf
`outputs('Sort_ByStatusPriority')[item()]?['Id']` zu — also in Status-Priority-
Reihenfolge. Die `TeilnehmerID` bleibt `add(item(), 1)`, d.h. Index+1.

Ergebnis: Angemeldete bekommen IDs 1..N, Warteliste bekommt IDs N+1..N+M — saubere,
lückenlose Sortierung. Beispiel mit 100 Plätzen: Wenn #98 (Angemeldet) abmeldet,
werden #99 → #98, #100 → #99, #101 (alter Warteliste-Erster, wird durch Nachrücken
auch Angemeldet) → #100, #102 (bleibt Warteliste) → #101 usw.

**Änderungen 2026-04-22: zusätzlich `Filter_Non_Waitlist`** (Query-Action) —
zählt Nicht-Warteliste-Items als Vorstufe zu `Count_Active`. Funktional identisch
zum alten Inline-`length(filter(...))`, aber lesbarer.

**Vorherige Änderungen 2026-04-20:**

1. `Get_Active_Participants` umbenannt zu `Get_Enrolled_Participants`. Filter geändert von
   `(Status eq 'Angemeldet') or (Status eq 'QR versendet') or (Status eq 'Eingecheckt')`
   auf `Status ne 'Abgemeldet'`. Vorher wurden Warteliste-Einträge beim Renummerieren
   übersprungen — führte zu Lücken in der TeilnehmerID-Sequenz. Jetzt bekommen Aktive
   + Warteliste gemeinsam fortlaufende IDs nach RegistrationDate.
2. Neuer Compose-Step `Count_Active` zählt Enrolled-Items mit `Status ≠ 'Warteliste'`.
   `Check_Nachrücken`-Bedingung vergleicht jetzt `Count_Active < MaxParticipants` statt
   `length(GenerateSPData) < MaxParticipants`. Nötig weil `GenerateSPData` nach Fix #1
   auch Warteliste-Einträge enthält und die alte Bedingung sonst nie mehr triggern würde.
3. `Promote_Waitlist.body` enthält nur noch `Status: Angemeldet` — KEIN `TeilnehmerID`
   mehr. Die korrekte TID wurde bereits im Batch-Update gesetzt (erste Warteliste in
   RegistrationDate-Reihenfolge = Count_Active + 1). Alte Logik `TID = length(GenerateSPData) + 1`
   hätte nach Fix #1 eine zu hohe TID erzeugt (= EnrolledCount + 1 statt Count_Active + 1).
4. `Queue_Email` (Nachrücker-Mail): Subject/Body/EventTitle-Spalte nutzen jetzt
   `first(outputs('Get_EventDetails')?['body/value'])?['Title']` statt
   `triggerOutputs()?['body/Title']` — letzteres war der DEX_IDReorder-Queue-Title
   ("Reorder: <EventName>"), nicht der echte Event-Name. Anrede + RecipientName nutzen
   jetzt `Vorname` statt `ParticipantName` (nur Vorname in der Begrüßung).
5. `Queue_Outlook.item/Title`: ebenfalls auf EventDetails.Title umgestellt — vorher
   "Einladen: Reorder: <Name>", jetzt "Einladen: <Name>".
6. Nachrücken-Template in `DEX_EmailTemplates` wird jetzt pre-wrapped gespeichert
   (komplettes Deloitte-Design inklusive Logo/Header/Footer), weil der PA-Flow den
   BodyHtml raw verwendet. Client-Code erkennt pre-wrapped Templates in
   `buildEmailFromTemplate()` und skippt den zweiten Wrap. App-seitig v5.33.0+ nötig.

```json
TRIGGER:
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

UPDATE_ITEM:
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
    "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "PatchItem" }
  },
  "runAfter": {}
}

SETTINGS:
{
  "type": "Compose",
  "inputs": {
    "siteAddress": "@{triggerOutputs()?['body/SubsiteUrl']}",
    "listName": "Teilnehmer",
    "batchSize": 250
  },
  "runAfter": { "Update_item": ["Succeeded"] }
}

GET_LISTITEMTYPE:
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "@outputs('Settings')?['siteAddress']",
      "parameters/method": "GET",
      "parameters/uri": "_api/web/lists/getbytitle('@{outputs('Settings')?['listName']}')?$select=ListItemEntityTypeFullName"
    },
    "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "HttpRequest" }
  },
  "runAfter": { "Settings": ["Succeeded"] }
}

GET_ENROLLED_PARTICIPANTS:
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "@outputs('Settings')?['siteAddress']",
      "parameters/method": "GET",
      "parameters/uri": "@concat('_api/web/lists/getbytitle(''', outputs('Settings')?['listName'], ''')/items?$select=Id,TeilnehmerID,Status,RegistrationDate&$filter=Status ne ''Abgemeldet''&$orderby=RegistrationDate asc&$top=5000')",
      "parameters/headers": { "Accept": "application/json;odata=nometadata" }
    },
    "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "HttpRequest" }
  },
  "runAfter": { "Get_ListItemType": ["Succeeded"] }
}

COUNT_ACTIVE:
{
  "type": "Compose",
  "inputs": "@length(filter(body('Get_Enrolled_Participants')?['value'], not(equals(item()?['Status'], 'Warteliste'))))",
  "runAfter": { "Get_Enrolled_Participants": ["Succeeded"] }
}

GENERATE_INDICES:
{
  "type": "Compose",
  "inputs": "@range(0, length(body('Get_Enrolled_Participants')?['value']))",
  "runAfter": { "Count_Active": ["Succeeded"] }
}

GENERATESPDATA:
{
  "type": "Select",
  "inputs": {
    "from": "@outputs('Generate_Indices')",
    "select": {
      "ID": "@body('Get_Enrolled_Participants')?['value'][item()]?['Id']",
      "TeilnehmerID": "@add(item(), 1)"
    }
  },
  "runAfter": { "Generate_Indices": ["Succeeded"] }
}

BATCHGUIDS:
{
  "type": "Compose",
  "inputs": { "batchGUID": "@{guid()}", "changeSetGUID": "@{guid()}" },
  "runAfter": { "GenerateSPData": ["Succeeded"] }
}

BATCHTEMPLATE:
{
  "type": "Compose",
  "inputs": "--changeset_@{outputs('BatchGuids')?['changeSetGUID']}\nContent-Type: application/http\nContent-Transfer-Encoding: binary\n\nPATCH @{outputs('Settings')?['siteAddress']}/_api/web/lists/getByTitle('@{outputs('Settings')?['listName']}')/items(|ID|) HTTP/1.1\nContent-Type: application/json;odata=verbose\nAccept: application/json;odata=verbose\nIf-Match: *\n\n{\"__metadata\":{\"type\":\"@{body('Get_ListItemType')?['d']?['ListItemEntityTypeFullName']}\"},\"TeilnehmerID\":|TID|}\n",
  "runAfter": { "BatchGuids": ["Succeeded"] }
}

PROCESS_BATCH_SCOPE:
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
              "parameters/headers": { "Content-Type": "multipart/mixed;boundary=batch_@{outputs('BatchGuids')?['batchGUID']}" },
              "parameters/body": "--batch_@{outputs('BatchGuids')?['batchGUID']}\nContent-Type: multipart/mixed; boundary=\"changeset_@{outputs('BatchGuids')?['changeSetGUID']}\"\nContent-Length: @{length(outputs('batchData'))}\nContent-Transfer-Encoding: binary\n\n@{outputs('batchData')}\n--changeset_@{outputs('BatchGuids')?['changeSetGUID']}--\n\n--batch_@{outputs('BatchGuids')?['batchGUID']}--\n"
            },
            "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "HttpRequest" }
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
        "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "GetItems" }
      },
      "runAfter": { "Loop_Batches": ["Succeeded"] }
    },
    "Check_Nachrücken": {
      "type": "If",
      "expression": {
        "and": [
          { "less": ["@outputs('Count_Active')", "@first(outputs('Get_EventDetails')?['body/value'])?['MaxParticipants']"] },
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
            "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "HttpRequest" }
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
                  "parameters/headers": { "Content-Type": "application/json;odata=verbose", "IF-MATCH": "*", "X-HTTP-Method": "MERGE", "Accept": "application/json;odata=verbose" },
                  "parameters/body": "{\"__metadata\":{\"type\":\"SP.Data.TeilnehmerListItem\"},\"Status\":\"Angemeldet\"}"
                },
                "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "HttpRequest" }
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
                "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "HttpRequest" }
              },
              "runAfter": { "Promote_Waitlist": ["Succeeded"] }
            },
            "Queue_Email": {
              "type": "OpenApiConnection",
              "inputs": {
                "parameters": {
                  "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                  "table": "57aa0840-df98-41ae-a39b-323c0b80ae3b",
                  "item/Title": "@replace(replace(coalesce(first(body('Get_Email_Template')?['d']?['results'])?['Subject'], concat('Spot available: ', first(outputs('Get_EventDetails')?['body/value'])?['Title'])), '{{Name}}', first(body('Get_Waitlist_First')?['d']?['results'])?['Vorname']), '{{EventTitle}}', first(outputs('Get_EventDetails')?['body/value'])?['Title'])",
                  "item/Recipient": "@first(body('Get_Waitlist_First')?['d']?['results'])?['ParticipantEmail']",
                  "item/RecipientName": "@first(body('Get_Waitlist_First')?['d']?['results'])?['Vorname']",
                  "item/EmailType/Value": "Nachruecken",
                  "item/EventTitle": "@first(outputs('Get_EventDetails')?['body/value'])?['Title']",
                  "item/Status/Value": "Pending",
                  "item/Body": "@replace(replace(coalesce(first(body('Get_Email_Template')?['d']?['results'])?['BodyHtml'], ''), '{{Name}}', first(body('Get_Waitlist_First')?['d']?['results'])?['Vorname']), '{{EventTitle}}', first(outputs('Get_EventDetails')?['body/value'])?['Title'])",
                  "item/EventId": "@triggerOutputs()?['body/EventId']"
                },
                "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "PostItem" }
              },
              "runAfter": { "Get_Email_Template": ["SUCCEEDED"] }
            },
            "Queue_Outlook": {
              "type": "OpenApiConnection",
              "inputs": {
                "parameters": {
                  "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                  "table": "d794655b-c950-416c-a478-5dbae285e46d",
                  "item/Title": "Einladen: @{first(outputs('Get_EventDetails')?['body/value'])?['Title']}",
                  "item/Attendee": "@first(body('Get_Waitlist_First')?['d']?['results'])?['ParticipantEmail']",
                  "item/EventId": "@triggerOutputs()?['body/EventId']",
                  "item/ActionType/Value": "Einladen",
                  "item/Status/Value": "Pending"
                },
                "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "PostItem" }
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
        "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "PatchItem" }
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
            "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "PatchItem" }
          }
        }
      },
      "runAfter": { "DEX_IDReorder": ["Failed"] }
    }
  },
  "runAfter": { "batchTemplate": ["Succeeded"] }
}
```

---

## 2. DEX_SEND_MAIL

**Trigger:** Neuer Eintrag in DEX_Emails
**Zweck:** E-Mails aus Queue versenden über Shared Mailbox (no_reply.events@deloitte.de)
**Letztes Update:** 2026-04-07

Ablauf: Trigger → Config laden (Logo + Default-Bild aus DEX_EmailTemplates via GetItems) → Event laden → Compose_Logo (aus Config) → Compose_Image (Event-Bild oder Default) → Platzhalter ersetzen → Email senden → Status=Sent

```json
TRIGGER:
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "57aa0840-df98-41ae-a39b-323c0b80ae3b"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline-1",
      "operationId": "GetOnNewItems"
    }
  },
  "recurrence": { "frequency": "Minute", "interval": 1 },
  "splitOn": "@triggerOutputs()?['body/value']"
}

GET_CONFIG (Logo + Default-Bild aus DEX_EmailTemplates via GetItems):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "2c428d35-e6fb-42f9-8a20-580acd6d05f4",
      "$filter": "TemplateType eq '_Config'",
      "$top": 1
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "GetItems"
    }
  },
  "runAfter": {}
}

GET_EVENT (Event-Daten für EventId):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "28457815-1163-4e92-8b08-3ae43f477d9e",
      "$filter": "@concat('ID eq ', triggerBody()?['EventId'])",
      "$top": 1
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "GetItems"
    }
  },
  "runAfter": { "Get_Config": ["Succeeded"] }
}

COMPOSE_LOGO (Logo Base64 aus Config):
{
  "type": "Compose",
  "inputs": "@first(body('Get_Config')?['value'])?['LogoBase64']",
  "runAfter": { "Get_Event": ["Succeeded"] }
}

COMPOSE_IMAGE (Event-Bild oder Default-Bild):
{
  "type": "Compose",
  "inputs": "@if(empty(first(outputs('Get_Event')?['body/value'])?['EmailImageBase64']), first(body('Get_Config')?['value'])?['DefaultImageBase64'], first(outputs('Get_Event')?['body/value'])?['EmailImageBase64'])",
  "runAfter": { "Compose_Logo": ["Succeeded"] }
}

SEND_EMAIL (Shared Mailbox):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "emailMessage/MailboxAddress": "no_reply.events@deloitte.de",
      "emailMessage/To": "@triggerBody()?['Recipient']",
      "emailMessage/Subject": "@triggerBody()?['Title']",
      "emailMessage/Body": "<p class=\"editor-paragraph\">@{replace(replace(triggerBody()?['Body'], '{{LOGO_URL}}', outputs('Compose_Logo')), '{{ORB_URL}}', outputs('Compose_Image'))}</p>",
      "emailMessage/Importance": "Normal"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
      "connection": "shared_office365",
      "operationId": "SharedMailboxSendEmailV2"
    }
  },
  "runAfter": { "Compose_Image": ["Succeeded"] }
}

SET_SENT:
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "57aa0840-df98-41ae-a39b-323c0b80ae3b",
      "id": "@triggerBody()?['ID']",
      "item/Title": "@triggerBody()?['Title']",
      "item/Status/Value": "Sent",
      "item/SentDate": "@utcNow()"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline-1",
      "operationId": "PatchItem"
    }
  },
  "runAfter": { "Send_an_email_from_a_shared_mailbox_(V2)": ["Succeeded"] }
}

SET_FAILED (Email-Versand fehlgeschlagen):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "57aa0840-df98-41ae-a39b-323c0b80ae3b",
      "id": "@triggerBody()?['ID']",
      "item/Title": "@triggerBody()?['Title']",
      "item/Status/Value": "Failed"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline-1",
      "operationId": "PatchItem"
    }
  },
  "runAfter": { "Send_an_email_from_a_shared_mailbox_(V2)": ["Failed"] }
}
```

---

## 3. DEX_CreateOutlookEvent

**Trigger:** Neuer Eintrag in DEX_Events
**Zweck:** Outlook-Kalendereintrag im Deloitte-Design erstellen (Logo + Event-Bild aus DEX_EmailTemplates) und iCalUId zurückschreiben
**Letztes Update:** 2026-04-09

Ablauf: Trigger (neues Event) → Config laden (Logo + Default-Bild) → Compose_Logo → Compose_Image → Platzhalter in OutlookBody ersetzen → Outlook-Termin mit HTML-Body erstellen (UTC-Zeit wird per `convertFromUtc` nach Europe/Berlin konvertiert) → CalendarLink in DEX_Events speichern

**Hinweis:** Der OutlookBody wird bereits in der SPFx-App im Deloitte-HTML-Template gewrappt (mit `{{LOGO_URL}}` und `{{ORB_URL}}` Platzhaltern). Der Flow ersetzt diese Platzhalter durch Base64-Bilder.

```json
TRIGGER (Neues Item in DEX_Events):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "28457815-1163-4e92-8b08-3ae43f477d9e"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "GetOnNewItems"
    }
  },
  "recurrence": { "frequency": "Minute", "interval": 1 },
  "splitOn": "@triggerOutputs()?['body/value']"
}

GET_CONFIG (Logo + Default-Bild aus DEX_EmailTemplates):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "parameters/method": "GET",
      "parameters/uri": "_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$select=LogoBase64,DefaultImageBase64&$top=1"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "HttpRequest"
    }
  },
  "runAfter": {}
}

COMPOSE_LOGO (Logo Base64 aus Config):
{
  "type": "Compose",
  "inputs": "@first(body('Get_Config')?['value'])?['LogoBase64']",
  "runAfter": { "Get_Config": ["SUCCEEDED"] }
}

COMPOSE_IMAGE (Event-Bild oder Default-Bild):
{
  "type": "Compose",
  "inputs": "@if(empty(triggerBody()?['EmailImageBase64']), first(body('Get_Config')?['value'])?['DefaultImageBase64'], triggerBody()?['EmailImageBase64'])",
  "runAfter": { "Compose_Logo": ["SUCCEEDED"] }
}

CREATE_EVENT_V4 (Outlook-Termin mit Deloitte-Design Body):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "table": "AAMkADU5YjlkMDBiLWU2MDktNGViMy1iNGIwLTI0YWFkNDkyN2VjMABGAAAAAABjJcNB5xJWS7D2nCeePixeBwAbtMj6YVUGQJroN6O--ImBAAAAAAEGAAAbtMj6YVUGQJroN6O--ImBAAKF4fCpAAA=",
      "item/subject": "@triggerBody()?['Title']",
      "item/start": "@convertFromUtc(triggerBody()?['StartDate'], 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss')",
      "item/end": "@convertFromUtc(triggerBody()?['EndDate'], 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss')",
      "item/timeZone": "(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna",
      "item/requiredAttendees": "@triggerBody()?['OrganizerEmail']",
      "item/body": "<p class=\"editor-paragraph\">@{replace(replace(coalesce(triggerBody()?['OutlookBody'], ''), '{{LOGO_URL}}', outputs('Compose_Logo')), '{{ORB_URL}}', outputs('Compose_Image'))}</p>"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
      "connection": "shared_office365",
      "operationId": "V4CalendarPostItem"
    }
  },
  "runAfter": { "Compose_Image": ["SUCCEEDED"] }
}

UPDATE_EVENT (CalendarLink zurückschreiben):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "28457815-1163-4e92-8b08-3ae43f477d9e",
      "id": "@triggerBody()?['ID']",
      "item/Title": "@triggerBody()?['Title']",
      "item/CalendarLink": "@outputs('Create_event_(V4)')?['body/iCalUId']"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "PatchItem"
    }
  },
  "runAfter": { "Create_event_(V4)": ["Succeeded"] }
}

SET_FAILED (Outlook-Termin Erstellung fehlgeschlagen):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "28457815-1163-4e92-8b08-3ae43f477d9e",
      "id": "@triggerBody()?['ID']",
      "item/Title": "@triggerBody()?['Title']",
      "item/OutlookEventId": "FAILED"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "PatchItem"
    }
  },
  "runAfter": { "Create_event_(V4)": ["Failed"] }
}
```

---

## 4. DEX_Outlook_Einladungen

**Trigger:** Neuer Eintrag in DEX_Outlook (alle 5 Minuten, Concurrency 1)
**Zweck:** Outlook-Termin verwalten: Teilnehmer einladen/ausladen, Event-Daten aktualisieren, Termin löschen (via Graph API).
**Letztes Update:** 2026-04-22 (v6.4, production): Sub-Event-Sonderlogik komplett entfernt. Sub-Events sind seit v6.4 eigene DEX_Events-Items mit gesetztem `ParentEventId` — sie laufen durch denselben Einladen/Ausladen/Update/Delete-Pfad wie Top-Level-Events, **keine** Override-Felder, **kein** separater Branch. Flow-Struktur entspricht wieder der ursprünglichen v5.x-Version.

### Flow-Struktur (v6.4)

```
Trigger (DEX_Outlook, alle 5 Min, Concurrency 1)
├── Is_DeleteEvent (ActionType == DeleteEvent, runAfter Trigger)
│   └── TRUE: Find_Outlook_Event_For_Delete (Graph GET by CalendarLink)
│             └── Outlook_Event_Found (length > 0)
│                 └── TRUE: Delete_Outlook_Event (Graph DELETE) → Set_Sent_DeleteEvent (SP Status=Sent)
├── Get_Event_Details (SharePoint Get Items DEX_Events by EventId, runAfter Is_DeleteEvent)
├── Init_RealEventId (string)
├── Init_Attendees (array)
└── Check_CalendarLink (Event.CalendarLink gesetzt?)
    ├── TRUE: Find_Outlook_Event (Graph GET by iCalUId)
    │         → Set_variable (varRealEventId)
    │         → Check_EventFound (varRealEventId nicht leer?)
    │           ├── TRUE: Get_Existing_Event → Set_variable_1 (var_Attendees)
    │           │         → Is_UpdateEvent (ActionType == UpdateEvent)
    │           │           ├── TRUE:  Build_Update_Body (Compose) → Send_an_HTTP_request (Graph PATCH Titel/Start/Ende/Body)
    │           │           └── FALSE: Check_ActionType (ActionType == Einladen)
    │           │                      ├── TRUE:  Add_Attendee + Update_Event_Einladen (PATCH attendees)
    │           │                      └── FALSE: Filter_Attendees + Update_Event_Ausladen (PATCH attendees)
    │           │         → Set_Sent (Status=Sent)
    │           └── FALSE: Set_Failed_1 (Event nicht in Outlook gefunden)
    └── FALSE: Set_Failed (kein CalendarLink im DEX_Events-Item)
```

### ActionTypes

- `Einladen` — einzelnen Teilnehmer zum Outlook-Termin hinzufügen
- `Ausladen` — einzelnen Teilnehmer aus dem Outlook-Termin entfernen
- `UpdateEvent` — Titel, Start, Ende **und Body** des Outlook-Termins aktualisieren (seit 2026-04-17: `OutlookBody` aus DEX_Events mit aufgelöstem `{{ORB_URL}}` per Graph PATCH; vorher blieb der Body vom initialen Create unverändert)
- `DeleteEvent` — kompletten Outlook-Termin löschen. Das DEX_Outlook-Queue-Item enthält `CalendarLink` (iCalUId) direkt, weil das zugehörige DEX_Events-Item bereits gelöscht ist wenn der Flow läuft.

### Warum keine Sub-Event-Sonderlogik (seit v6.4)

Sub-Events sind eigene DEX_Events-Items mit `ParentEventId` ≠ leer. Dadurch:
- Der **`DEX_CreateOutlookEvent`**-Flow legt für jedes Sub-Event automatisch einen eigenen Kalendertermin an (wie für Top-Level-Events).
- Anmeldungen werden in der **eigenen Subsite** des Sub-Events gespeichert (eigene Teilnehmerliste, QR-Codes, Warteliste).
- **`DEX_Outlook_Einladungen`** (dieser Flow) bekommt für jede Sub-Event-Anmeldung eine ganz normale `Einladen`-Queue-Zeile mit EventId=Sub-Event-ID. Er findet das Calendar-Event über den CalendarLink im Sub-Event-DEX_Events-Item und fügt den Attendee hinzu — keine Sonderwege nötig.
- Bei **Parent-Löschung** löscht die App kaskadierend alle Child-Events (`EventContext.deleteEvent` iteriert `events.filter(e.parentEventId === parentId)`). Jedes Child-Delete queued sein eigenes `DeleteEvent`-Item → Is_DeleteEvent-Branch räumt den Kalender auf.

### Parent-Event-Update (seit 2026-04-17, unverändert)

Ablauf:
1. Trigger (neues DEX_Outlook-Item)
2. **Is_DeleteEvent?** → Ja: Outlook-Termin per `triggerBody()?['CalendarLink']` finden (iCalUId) → DELETE → Status=Sent. Nein: weiter.
3. Event-Details laden (DEX_Events via EventId) → CalendarLink vorhanden? → Outlook-Event per iCalUId finden → Event-ID speichern → Event gefunden?
4. Bestehende Attendees laden → Is_UpdateEvent? → Ja: PATCH Titel/Start/Ende **+ Body** → Nein: Einladen/Ausladen → Status=Sent

**ActionTypes:**
- `Einladen` — einzelnen Teilnehmer zum Outlook-Termin hinzufuegen
- `Ausladen` — einzelnen Teilnehmer aus dem Outlook-Termin entfernen
- `UpdateEvent` — Titel, Start, Ende **und Body** des Outlook-Termins aktualisieren
  (seit 2026-04-17 wird `OutlookBody` aus DEX_Events mit aufgeloestem
  `{{ORB_URL}}` per Graph PATCH gesetzt — vorher blieb der Body vom
  initialen Create unveraendert)
- `DeleteEvent` — kompletten Outlook-Termin loeschen. Das DEX_Outlook-Queue-Item
  enthaelt `CalendarLink` (iCalUId) direkt, weil das zugehoerige DEX_Events-Item
  bereits geloescht ist, wenn der Flow laeuft.

**Concurrency:** 1 (sequentielle Verarbeitung, max 100 wartende Runs)

**UpdateEvent-Pattern** (seit 2026-04-17): String-Concat mit `json()` ist nicht
robust genug fuer beliebige HTML-Inhalte (Quotes/Newlines/Sonderzeichen brechen
das Parsing). Stattdessen wird der Body via Compose-Action vorgebaut und im
HTTP-PATCH referenziert — Logic Apps escaped die `@{...}`-Tokens automatisch.

**Compose `Build_Update_Body`** (vor `Send_an_HTTP_request` ausfuehren):
```json
{
  "subject": "@{first(outputs('Get_Event_Details')?['body/value'])?['Title']}",
  "start": {
    "dateTime": "@{convertFromUtc(first(outputs('Get_Event_Details')?['body/value'])?['StartDate'], 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss')}",
    "timeZone": "W. Europe Standard Time"
  },
  "end": {
    "dateTime": "@{convertFromUtc(first(outputs('Get_Event_Details')?['body/value'])?['EndDate'], 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss')}",
    "timeZone": "W. Europe Standard Time"
  },
  "body": {
    "contentType": "html",
    "content": "@{replace(coalesce(first(outputs('Get_Event_Details')?['body/value'])?['OutlookBody'], ''), '{{ORB_URL}}', coalesce(first(outputs('Get_Event_Details')?['body/value'])?['EmailImageBase64'], ''))}"
  }
}
```

**`Send_an_HTTP_request`-Body** (PATCH zur Graph API):
```
@outputs('Build_Update_Body')
```

### Is_DeleteEvent-Branch (ganz am Anfang, vor Get_Event_Details)

```json
{
  "type": "If",
  "expression": {
    "and": [
      { "equals": ["@triggerBody()?['ActionType']?['Value']", "DeleteEvent"] }
    ]
  },
  "actions": {
    "Find_Outlook_Event_For_Delete": {
      "type": "OpenApiConnection",
      "inputs": {
        "parameters": {
          "Uri": "@concat('https://graph.microsoft.com/v1.0/users/no_reply.events@deloitte.de/events?$filter=iCalUId eq ''', triggerBody()?['CalendarLink'], '''')",
          "Method": "GET",
          "ContentType": "application/json"
        },
        "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365", "connection": "shared_office365", "operationId": "HttpRequest" }
      }
    },
    "Outlook_Event_Found": {
      "type": "If",
      "expression": { "and": [ { "greater": ["@length(body('Find_Outlook_Event_For_Delete')?['value'])", 0] } ] },
      "actions": {
        "Delete_Outlook_Event": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "Uri": "@concat('https://graph.microsoft.com/v1.0/users/no_reply.events@deloitte.de/events/', first(body('Find_Outlook_Event_For_Delete')?['value'])?['id'])",
              "Method": "DELETE",
              "ContentType": "application/json"
            },
            "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365", "connection": "shared_office365", "operationId": "HttpRequest" }
          }
        },
        "Set_Sent_DeleteEvent": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
              "table": "d794655b-c950-416c-a478-5dbae285e46d",
              "id": "@triggerBody()?['ID']",
              "item/Title": "@triggerBody()?['Title']",
              "item/Status/Value": "Sent",
              "item/SentDate": "@utcNow()"
            },
            "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "PatchItem" }
          },
          "runAfter": { "Delete_Outlook_Event": ["Succeeded", "Failed"] }
        }
      },
      "else": { "actions": {} },
      "runAfter": { "Find_Outlook_Event_For_Delete": ["Succeeded"] }
    }
  },
  "else": { "actions": {} },
  "runAfter": {}
}
```

Wichtig: `Get_Event_Details` hat nach Einfuegen dieser Condition `runAfter = { "Is_DeleteEvent": ["Succeeded"] }`, damit die Haupt-Logik nur laeuft wenn es kein DeleteEvent ist. Der Else-Zweig von `Is_DeleteEvent` ist leer (alle weiteren Actions kommen sowieso nach der Condition).

### SharePoint-Liste DEX_Outlook

Fuer `DeleteEvent` muessen folgende Schema-Aenderungen vorgenommen werden (werden bei neuen Listen automatisch von `ensureOutlookList()` angelegt, bei bestehenden muss der Admin sie manuell ergaenzen):

- **Choice `ActionType`** erweitern um `DeleteEvent`
- **Neue Spalte** `CalendarLink` (Multiple lines of text, plain) — enthaelt die iCalUId, damit der Flow das Outlook-Event auch ohne Zugriff auf DEX_Events finden kann.


```json
TRIGGER (Neues Item in DEX_Outlook, Concurrency: 1):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "d794655b-c950-416c-a478-5dbae285e46d"
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

GET_EVENT_DETAILS (Event-Daten für EventId):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "28457815-1163-4e92-8b08-3ae43f477d9e",
      "$filter": "@concat('ID eq ', triggerBody()?['EventId'])",
      "$top": 1
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "GetItems"
    }
  },
  "runAfter": {}
}

INIT_REALEVENTID (Variable für Outlook Event-ID):
{
  "type": "InitializeVariable",
  "inputs": { "variables": [{ "name": "varRealEventId", "type": "string" }] },
  "runAfter": { "Get_Event_Details": ["Succeeded"] }
}

INIT_ATTENDEES (Variable für Teilnehmer-Array):
{
  "type": "InitializeVariable",
  "inputs": { "variables": [{ "name": "var_Attendees", "type": "array" }] },
  "runAfter": { "Init_RealEventId": ["Succeeded"] }
}

CHECK_CALENDARLINK (CalendarLink vorhanden? + Einladen/Ausladen/UpdateEvent):
{
  "type": "If",
  "expression": {
    "and": [{ "greater": ["@length(coalesce(first(outputs('Get_Event_Details')?['body/value'])?['CalendarLink'], ''))", 0] }]
  },
  "actions": {
    "Set_Sent": {
      "type": "OpenApiConnection",
      "inputs": {
        "parameters": {
          "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
          "table": "d794655b-c950-416c-a478-5dbae285e46d",
          "id": "@triggerBody()?['ID']",
          "item/Title": "@triggerBody()?['Title']",
          "item/Status/Value": "Sent"
        },
        "host": {
          "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
          "connection": "shared_sharepointonline",
          "operationId": "PatchItem"
        }
      },
      "runAfter": { "Check_EventFound": ["Succeeded"] }
    },
    "Set_variable": {
      "type": "SetVariable",
      "inputs": { "name": "varRealEventId", "value": "@first(body('Find_Outlook_Event')?['value'])?['id']" },
      "runAfter": { "Find_Outlook_Event": ["Succeeded"] }
    },
    "Find_Outlook_Event": {
      "type": "OpenApiConnection",
      "inputs": {
        "parameters": {
          "Uri": "@concat('https://graph.microsoft.com/v1.0/users/no_reply.events@deloitte.de/events?$filter=iCalUId eq ''', first(outputs('Get_Event_Details')?['body/value'])?['CalendarLink'], '''')",
          "Method": "GET",
          "ContentType": "application/json"
        },
        "host": {
          "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
          "connection": "shared_office365",
          "operationId": "HttpRequest"
        }
      }
    },
    "Check_EventFound": {
      "type": "If",
      "expression": {
        "and": [{ "greater": ["@length(coalesce(variables('varRealEventId'), ''))", 0] }]
      },
      "actions": {
        "Get_Existing_Event": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "Uri": "@concat('https://graph.microsoft.com/v1.0/users/no_reply.events@deloitte.de/events/', variables('varRealEventId'))",
              "Method": "GET",
              "ContentType": "application/json"
            },
            "host": {
              "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
              "connection": "shared_office365",
              "operationId": "HttpRequest"
            }
          }
        },
        "Set_variable_1": {
          "type": "SetVariable",
          "inputs": { "name": "var_Attendees", "value": "@body('Get_Existing_Event')?['attendees']" },
          "runAfter": { "Get_Existing_Event": ["Succeeded"] }
        },
        "Set_Failed_1_1": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
              "table": "d794655b-c950-416c-a478-5dbae285e46d",
              "id": "@triggerBody()?['ID']",
              "item/Title": "@triggerBody()?['Title']",
              "item/Status/Value": "Failed"
            },
            "host": {
              "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
              "connection": "shared_sharepointonline",
              "operationId": "PatchItem"
            }
          },
          "runAfter": { "Is_UpdateEvent": ["Failed"] }
        },
        "Is_UpdateEvent": {
          "type": "If",
          "expression": {
            "and": [{ "equals": ["@triggerBody()?['ActionType']?['Value']", "UpdateEvent"] }]
          },
          "actions": {
            "Send_an_HTTP_request (PATCH Event-Daten)": {
              "type": "OpenApiConnection",
              "inputs": {
                "parameters": {
                  "Uri": "@concat('https://graph.microsoft.com/v1.0/users/no_reply.events@deloitte.de/events/', variables('varRealEventId'))",
                  "Method": "PATCH",
                  "Body": "@json(concat('{\"subject\":\"', first(outputs('Get_Event_Details')?['body/value'])?['Title'], '\",\"start\":{\"dateTime\":\"', convertFromUtc(first(outputs('Get_Event_Details')?['body/value'])?['StartDate'], 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss'), '\",\"timeZone\":\"W. Europe Standard Time\"},\"end\":{\"dateTime\":\"', convertFromUtc(first(outputs('Get_Event_Details')?['body/value'])?['EndDate'], 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss'), '\",\"timeZone\":\"W. Europe Standard Time\"}}'))",
                  "ContentType": "application/json"
                },
                "host": {
                  "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                  "connection": "shared_office365",
                  "operationId": "HttpRequest"
                }
              }
            }
          },
          "else": {
            "actions": {
              "Check_ActionType (Einladen oder Ausladen)": {
                "type": "If",
                "expression": {
                  "and": [{ "equals": ["@triggerBody()?['ActionType']?['Value']", "Einladen"] }]
                },
                "actions": {
                  "Add_Attendee": {
                    "type": "AppendToArrayVariable",
                    "inputs": {
                      "name": "var_Attendees",
                      "value": "@json(concat('{\"type\":\"required\",\"status\":{\"response\":\"none\",\"time\":\"0001-01-01T00:00:00Z\"},\"emailAddress\":{\"name\":\"', triggerBody()?['Attendee'], '\",\"address\":\"', triggerBody()?['Attendee'], '\"}}'))"
                    }
                  },
                  "Update_Event_Einladen": {
                    "type": "OpenApiConnection",
                    "inputs": {
                      "parameters": {
                        "Uri": "@concat('https://graph.microsoft.com/v1.0/users/no_reply.events@deloitte.de/events/', variables('varRealEventId'))",
                        "Method": "PATCH",
                        "Body": "@json(concat('{\"attendees\":', string(variables('var_Attendees')), '}'))",
                        "ContentType": "application/json"
                      },
                      "host": {
                        "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                        "connection": "shared_office365",
                        "operationId": "HttpRequest"
                      }
                    },
                    "runAfter": { "Add_Attendee": ["Succeeded"] }
                  }
                },
                "else": {
                  "actions": {
                    "Filter_Attendees": {
                      "type": "Query",
                      "inputs": {
                        "from": "@variables('var_Attendees')",
                        "where": "@not(equals(toLower(item()?['emailAddress']?['address']),toLower(triggerBody()?['Attendee'])))"
                      }
                    },
                    "Update_Event_Ausladen": {
                      "type": "OpenApiConnection",
                      "inputs": {
                        "parameters": {
                          "Uri": "@concat('https://graph.microsoft.com/v1.0/users/no_reply.events@deloitte.de/events/', variables('varRealEventId'))",
                          "Method": "PATCH",
                          "Body": "@json(concat('{\"attendees\":', string(body('Filter_Attendees')), '}'))",
                          "ContentType": "application/json"
                        },
                        "host": {
                          "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                          "connection": "shared_office365",
                          "operationId": "HttpRequest"
                        }
                      },
                      "runAfter": { "Filter_Attendees": ["Succeeded"] }
                    }
                  }
                }
              }
            }
          },
          "runAfter": { "Set_variable_1": ["Succeeded"] }
        }
      },
      "else": {
        "actions": {
          "Set_Failed_1 (Event nicht in Outlook gefunden)": {
            "type": "OpenApiConnection",
            "inputs": {
              "parameters": {
                "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                "table": "d794655b-c950-416c-a478-5dbae285e46d",
                "id": "@triggerBody()?['ID']",
                "item/Title": "@triggerBody()?['Title']",
                "item/Status/Value": "Failed"
              },
              "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                "connection": "shared_sharepointonline",
                "operationId": "PatchItem"
              }
            }
          }
        }
      },
      "runAfter": { "Set_variable": ["Succeeded"] }
    }
  },
  "else": {
    "actions": {
      "Set_Failed (Kein CalendarLink)": {
        "type": "OpenApiConnection",
        "inputs": {
          "parameters": {
            "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
            "table": "d794655b-c950-416c-a478-5dbae285e46d",
            "id": "@triggerBody()?['ID']",
            "item/Title": "@triggerBody()?['Title']",
            "item/Status/Value": "Failed"
          },
          "host": {
            "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
            "connection": "shared_sharepointonline",
            "operationId": "PatchItem"
          }
        }
      }
    }
  },
  "runAfter": { "Init_Attendees": ["Succeeded"] }
}
```

---

## Listen-GUIDs (aktuell, Stand 2026-04-07)

| Liste | GUID |
|-------|------|
| DEX_IDReorder | 9d46ff77-5fe2-4e1d-9b93-14b9dca1a360 |
| DEX_Events | 28457815-1163-4e92-8b08-3ae43f477d9e |
| DEX_Emails | 57aa0840-df98-41ae-a39b-323c0b80ae3b |
| DEX_Outlook | d794655b-c950-416c-a478-5dbae285e46d |
| DEX_EmailTemplates | 2c428d35-e6fb-42f9-8a20-580acd6d05f4 |

---

# FLOW: DEX_OutlookDeclineHandler (Mail-basiert, Stand 2026-04-14)

**Zweck:** Wenn ein User einen Outlook-Kalender-Termin im `no_reply.events@deloitte.de`-Postfach **ablehnt**, aber in der DEX-Teilnehmerliste noch als angemeldet steht, bekommt er eine Erinnerungsmail mit einem Action-Button "Anmeldung stornieren". Der Link öffnet die DEX App im Abmelde-Modus (Deep-Link `?action=cancel&event=<eventNumber>`).

**Trigger:** `When a new email arrives (V3)` auf der Shared Mailbox — nicht `When an event is modified`, weil der Event-Trigger keine pro-Attendee-`status.response`-Felder liefert und damit nicht erkennbar ist, WER abgelehnt hat.

## Bekannte Einschränkungen

| Fall | Verhalten | Abdeckung |
|------|-----------|-----------|
| User lehnt ab + sendet Antwort (Default) | ✅ Mail kommt → Flow feuert → Reminder wird gequeued | ~85% |
| User lehnt ab + "keine Antwort senden" | ❌ Keine Mail → kein Reminder (silent decline) | Lücke |
| User hat Outlook auf DE/EN/FR/IT | ✅ Subject-Filter deckt alle ab | OK |
| User hat Outlook auf anderer Sprache (PL/TR/ES/...) | ❌ Subject wird nicht erkannt | Lücke |
| Zwei Events mit identischem Titel | ⚠️ Flow nimmt das erste in DEX_Events | Edge Case |

Für 100% Abdeckung wäre ein Graph-API-Polling-Flow nötig (kein Ziel für jetzt).

## UI-Anleitung zum Anlegen (Schritt für Schritt)

### 1. Neuer Cloud-Flow anlegen

1. https://make.powerautomate.com öffnen.
2. Links **+ Create** → **Automated cloud flow**.
3. **Flow name:** `DEX_OutlookDeclineHandler`.
4. Trigger: `When a new email arrives V3` eintippen → **Office 365 Outlook — When a new email arrives (V3)** auswählen.
5. **Create**.

### 2. Trigger konfigurieren

- **Folder:** `Inbox`.
- **Show advanced options** klicken:
  - **Original Mailbox Address:** `no_reply.events@deloitte.de`.
  - **Include Attachments:** `No`.
  - **Subject Filter:** leer lassen (Sprachabhängigkeit wird in Schritt 3 über Condition gelöst).
  - **Importance, From, To, CC, Has Attachment:** leer lassen.

### 3. Condition: Ist das eine Decline-Mail? (`Is_Decline_Mail`)

- **+ New step** → **Control — Condition**.
- Linke Seite: **Expression-Tab (fx)** →
  ```
  or(
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'declined:'),
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'declined '),
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'abgelehnt:'),
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'abgelehnt '),
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'refusé'),
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'rifiutato:'),
    and(
      or(
        startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'fw:'),
        startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'wg:')
      ),
      or(
        contains(toLower(coalesce(triggerOutputs()?['body/bodyPreview'], '')), 'declined:'),
        contains(toLower(coalesce(triggerOutputs()?['body/bodyPreview'], '')), 'abgelehnt:')
      )
    )
  )
  ```
  Der letzte `and(...)`-Block deckt **weitergeleitete** Decline-Mails ab:
  Subject startet mit `FW:` / `WG:` UND der `bodyPreview` enthaelt `Declined:`
  oder `Abgelehnt:` (typisches Outlook-Forward-Format).
- Operator: `is equal to`.
- Rechte Seite: **Expression (fx)** → `true`.
- Rename **(⋮)** → `Is_Decline_Mail`.

Alle weiteren Schritte im **If yes**-Zweig; **If no** bleibt leer.

### 4. `Cleaned_Subject` (Compose)

- **Data Operation — Compose**.
- **Inputs (fx):**
  ```
  trim(
    replace(
      replace(
        replace(
          replace(
            replace(
              replace(
                replace(
                  replace(triggerOutputs()?['body/subject'], 'FW:', ''),
                  'WG:', ''),
                'Declined:', ''),
              'Declined ', ''),
            'Abgelehnt:', ''),
          'Abgelehnt ', ''),
        'Refusé :', ''),
      'Rifiutato:', '')
  )
  ```
  Zusaetzlich zu den sechs Decline-Prefixen werden auch `FW:` und `WG:`
  abgeschnitten, damit weitergeleitete Decline-Mails den reinen Event-Titel
  in `Cleaned_Subject` haben.
- Rename → `Cleaned_Subject`.

### 4a. `Real_Sender` (Compose)

- **Data Operation — Compose** (NACH `Cleaned_Subject`).
- **Inputs (fx):**
  ```
  if(or(startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'fw:'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'wg:')), trim(first(split(first(skip(split(coalesce(triggerOutputs()?['body/bodyPreview'], ''), '<'), 1)), '>'))), triggerOutputs()?['body/from'])
  ```
  Fuer direkte Decline-Mails = `body/from` wie bisher. Fuer Forwards extrahiert
  die Expression die erste `<email>`-Adresse aus dem `bodyPreview`. Bei
  "On Behalf Of"-Forwards ist das die Adresse der **Assistenz**, nicht des
  Principals — darum braucht's unten zusaetzlich die Name-basierte Suche.
- Rename → `Real_Sender`.

### 4b. `Decliner_Lastname` (Compose)

- **Data Operation — Compose** (NACH `Real_Sender`).
- **Inputs (fx):**
  ```
  if(contains(coalesce(triggerOutputs()?['body/bodyPreview'], ''), 'On Behalf Of '), trim(first(split(last(split(triggerOutputs()?['body/bodyPreview'], 'On Behalf Of ')), ','))), '')
  ```
  Extrahiert den Nachnamen aus Outlook's "On Behalf Of <Nachname>, <Titel>
  <Vorname>"-Pattern. Leer bei direkten Decline-Mails.
- Rename → `Decliner_Lastname`.

### 4c. `Decliner_Firstname` (Compose)

- **Data Operation — Compose** (NACH `Decliner_Lastname`).
- **Inputs (fx):**
  ```
  if(contains(coalesce(triggerOutputs()?['body/bodyPreview'], ''), 'On Behalf Of '), trim(replace(replace(replace(last(split(first(split(last(split(triggerOutputs()?['body/bodyPreview'], 'On Behalf Of ')), decodeUriComponent('%0D%0A'))), ',')), 'Prof. ', ''), 'Dr. ', ''), 'Dipl.-Ing. ', '')), '')
  ```
  Extrahiert den Vornamen nach dem Komma, strippt die haeufigsten Titel
  (`Prof. `, `Dr. `, `Dipl.-Ing. `). Beispiel: `"Nibler, Dr. Marcus"` →
  `"Marcus"`. Leer bei direkten Decline-Mails.
- Rename → `Decliner_Firstname`.

### 5. `Get_DEX_Event` (SharePoint Get items)

- **SharePoint — Get items**.
- **Site Address:** `DOL-c-DE-EventExperiencePlatform`.
- **List Name:** `DEX_Events`.
- **Show advanced options** → **Filter Query (fx):**
  ```
  concat('Title eq ''', replace(outputs('Cleaned_Subject'), '''', ''''''), '''')
  ```
- **Top Count:** `1`.
- **Configure run after** → `Decliner_Firstname` → `Succeeded`.
- Rename → `Get_DEX_Event`.

### 6. Condition `Event_Found`

- Linke Seite **(fx):** `length(outputs('Get_DEX_Event')?['body/value'])`
- Operator: `is greater than`
- Rechte Seite: `0` (Plain Text).
- Rename → `Event_Found`.

### 7. `Get_Teilnehmer_Entry` (Send HTTP request to SharePoint, im If yes)

- **SharePoint — Send an HTTP request to SharePoint**.
- **Site Address (fx):** `first(outputs('Get_DEX_Event')?['body/value'])?['SubsiteUrl']`
- **Method:** `GET`.
- **Uri (fx):** Konditional, filtert je nach verfuegbaren Daten nach E-Mail
  oder Name:
  ```
  concat('_api/web/lists/getbytitle(''Teilnehmer'')/items?$filter=', if(empty(outputs('Decliner_Lastname')), concat('ParticipantEmail eq ''', replace(outputs('Real_Sender'), '''', ''''''), ''''), if(empty(outputs('Decliner_Firstname')), concat('Nachname eq ''', replace(outputs('Decliner_Lastname'), '''', ''''''), ''''), concat('Vorname eq ''', replace(outputs('Decliner_Firstname'), '''', ''''''), ''' and Nachname eq ''', replace(outputs('Decliner_Lastname'), '''', ''''''), ''''))), ' and Status ne ''Abgemeldet''&$top=1&$select=Id,Status,Vorname,Nachname,ParticipantEmail')
  ```
  Drei Faelle:
  1. **Kein OnBehalfOf** → `ParticipantEmail eq '<Real_Sender>'` (direkter
     Decline oder regulaerer Forward)
  2. **OnBehalfOf mit Vor- und Nachname** → `Vorname eq '<x>' and Nachname eq '<y>'`
     (praeziseste Variante, vermeidet Kollision bei mehreren Niblers im Event)
  3. **OnBehalfOf nur Nachname** (Vorname-Parsing gescheitert) →
     `Nachname eq '<y>'` (Fallback)

  **Wichtig:** `ParticipantEmail` MUSS mit im `$select` stehen, damit
  `Final_Recipient_Email` die echte E-Mail-Adresse des Principals (z.B.
  Dr. Nibler) laden kann — `Real_Sender` ist bei OnBehalfOf-Faellen die
  Assistenz-Adresse und darf NICHT als Reminder-Empfaenger verwendet werden.
- **Headers:**
  - `Accept: application/json;odata=nometadata`
- Rename → `Get_Teilnehmer_Entry`.

### 7a. `Final_Recipient_Email` (Compose)

- **Data Operation — Compose** (NACH `Get_Teilnehmer_Entry`, VOR
  `Still_Registered`).
- **Inputs (fx):**
  ```
  coalesce(first(body('Get_Teilnehmer_Entry')?['value'])?['ParticipantEmail'], outputs('Real_Sender'))
  ```
  Nimmt die ParticipantEmail aus dem gefundenen Teilnehmer-Eintrag. Fallback
  auf `Real_Sender` wenn kein Teilnehmer gefunden (dann geht der Reminder eh
  nicht raus, weil `Still_Registered=false`, aber verhindert null-errors).
- **Configure run after** → `Get_Teilnehmer_Entry` → `Succeeded`.
- Rename → `Final_Recipient_Email`.

### 8. Condition `Still_Registered`

- Linke Seite **(fx):** `length(body('Get_Teilnehmer_Entry')?['value'])`
- Operator: `is greater than`
- Rechte Seite: `0`.
- **Configure run after** → `Final_Recipient_Email` → `Succeeded`.
- Rename → `Still_Registered`.

### 9. `Get_Existing_Reminder` (SharePoint Get items, im Still_Registered/yes)

- **SharePoint — Get items**.
- **Site Address:** `DOL-c-DE-EventExperiencePlatform`.
- **List Name:** `DEX_Emails`.
- **Filter Query (fx):**
  ```
  concat(
    'EmailType eq ''OutlookDeclineReminder'' and EventId eq ''',
    first(outputs('Get_DEX_Event')?['body/value'])?['ID'],
    ''''
  )
  ```
  Kein `Recipient`-Filter im OData (Multi-Line-Note-Feld, `eq` nicht moeglich)
  — der nachgeschaltete `Filter_By_Recipient`-Schritt pickt den aktuellen
  Sender heraus.
- **Top Count:** `20`.
- Rename → `Get_Existing_Reminder`.

### 9a. `Filter_By_Recipient` (Data Operation — Filter array)

- **Data Operation — Filter array**.
- **From (fx):** `body('Get_Existing_Reminder')?['value']`
- **Condition (fx):**
  ```
  contains(concat(';', replace(item()?['Recipient'], ' ', ''), ';'), concat(';', outputs('Final_Recipient_Email'), ';'))
  ```
  Vergleicht gegen `outputs('Final_Recipient_Email')` (die echte Empfaenger-
  Adresse aus dem Teilnehmer-Eintrag), damit bei OnBehalfOf-Forwards die
  bereits versendete Reminder-Mail korrekt gefunden wird.
- Rename → `Filter_By_Recipient`.

### 10. Condition `No_Reminder_Yet`

- Linke Seite **(fx):** `length(body('Filter_By_Recipient'))`
- Operator: `is equal to`
- Rechte Seite: `0`.
- Rename → `No_Reminder_Yet`.

### 11. `Get_Reminder_Template` (SharePoint Get items, im No_Reminder_Yet/yes)

- **SharePoint — Get items**.
- **Site Address:** `DOL-c-DE-EventExperiencePlatform`.
- **List Name:** `DEX_EmailTemplates`.
- **Filter Query (fx):** Konditional — Direkt-Decliner bekommen die schlichte
  Variante, OnBehalfOf bekommen die Variante mit Assistant-Forward-Button:
  ```
  concat('TemplateType eq ''', if(empty(outputs('Decliner_Lastname')), 'OutlookDeclineReminder', 'OutlookDeclineReminder_OnBehalfOf'), ''' and Language eq ''', coalesce(first(outputs('Get_DEX_Event')?['body/value'])?['EmailLanguage'], 'EN'), '''')
  ```
- **Top Count:** `1`.
- Rename → `Get_Reminder_Template`.

### 11a. `Assistant_Forward_Mailto` (Compose, nach `Get_Reminder_Template`)

- **Data Operation — Compose**.
- **Inputs (fx):** Baut die `mailto:`-URL fuer den Assistant-Forward-Button im
  OnBehalfOf-Template — vorausgefuellt mit Event-Organizer-Adressen, Subject
  und Body, der den Partner-Namen + Event-Titel enthaelt:
  ```
  concat('mailto:', first(outputs('Get_DEX_Event')?['body/value'])?['OrganizerEmail'], '?subject=', encodeUriComponent(concat('Bitte um Abmeldung: ', first(body('Get_Teilnehmer_Entry')?['value'])?['Vorname'], ' ', first(body('Get_Teilnehmer_Entry')?['value'])?['Nachname'], ' — ', first(outputs('Get_DEX_Event')?['body/value'])?['Title'])), '&body=', encodeUriComponent(concat('Hallo,', decodeUriComponent('%0D%0A%0D%0A'), 'ich (Assistenz) habe in Vertretung den Outlook-Termin für ', first(body('Get_Teilnehmer_Entry')?['value'])?['Vorname'], ' ', first(body('Get_Teilnehmer_Entry')?['value'])?['Nachname'], ' (', outputs('Final_Recipient_Email'), ') abgelehnt.', decodeUriComponent('%0D%0A%0D%0A'), 'Bitte storniere die Anmeldung für die Person für das Event "', first(outputs('Get_DEX_Event')?['body/value'])?['Title'], '" über das Admin Center der Event Experience Platform.', decodeUriComponent('%0D%0A%0D%0A'), 'Danke!')))
  ```
  Das Standard-Decliner-Template enthaelt keinen `{{AssistantForwardUrl}}`-
  Platzhalter, dort wird der Compose-Output einfach nicht verwendet.
- **Configure run after** → `Get_Reminder_Template` → `Succeeded`.
- Rename → `Assistant_Forward_Mailto`.

### 12. `Create_Reminder_Queue_Item` (SharePoint Create item, nach `Get_Reminder_Template`)

- **SharePoint — Create item**.
- **Site Address:** `DOL-c-DE-EventExperiencePlatform`.
- **List Name:** `DEX_Emails`.
- **Title (fx):**
  ```
  replace(coalesce(first(outputs('Get_Reminder_Template')?['body/value'])?['Subject'], concat('Outlook-Abmeldung-Reminder: ', first(outputs('Get_DEX_Event')?['body/value'])?['Title'])), '{{EventTitle}}', first(outputs('Get_DEX_Event')?['body/value'])?['Title'])
  ```
- **Recipient (fx):** `outputs('Final_Recipient_Email')`
- **RecipientName (fx):**
  ```
  coalesce(first(body('Get_Teilnehmer_Entry')?['value'])?['Vorname'], outputs('Final_Recipient_Email'))
  ```
  Vorname aus dem gefundenen Teilnehmer-Eintrag, damit die Mail mit
  "Dear Marcus," statt "Dear mparschalk@deloitte.de," beginnt. Fallback auf
  `Final_Recipient_Email` falls Vorname-Feld leer ist.
- **EmailType Value:** `OutlookDeclineReminder`.
- **EventTitle (fx):** `first(outputs('Get_DEX_Event')?['body/value'])?['Title']`
- **EventId (fx):** `first(outputs('Get_DEX_Event')?['body/value'])?['ID']`
- **Status Value:** `Pending`.
- **Body (fx):** Template laden und vier Platzhalter ersetzen — `DEX_SEND_MAIL`
  ersetzt danach nur noch `{{LOGO_URL}}` / `{{ORB_URL}}`:
  ```
  replace(replace(replace(replace(coalesce(first(outputs('Get_Reminder_Template')?['body/value'])?['BodyHtml'], ''), '{{Name}}', coalesce(first(body('Get_Teilnehmer_Entry')?['value'])?['Vorname'], outputs('Final_Recipient_Email'))), '{{EventTitle}}', first(outputs('Get_DEX_Event')?['body/value'])?['Title']), '{{CancelUrl}}', concat('https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView&action=cancel&event=', string(first(outputs('Get_DEX_Event')?['body/value'])?['EventNumber']))), '{{AssistantForwardUrl}}', outputs('Assistant_Forward_Mailto'))
  ```
  **Wichtig beim `{{Name}}`-Replace:** `coalesce(...Vorname, Final_Recipient_Email)`
  nutzen (nicht `Real_Sender`), sonst landet bei OnBehalfOf-Faellen die
  Assistenz-Mail-Adresse in der Anrede statt des Principal-Vornamens. Der
  vierte Replace `{{AssistantForwardUrl}}` ist NUR im OnBehalfOf-Template als
  Platzhalter vorhanden — bei Direkt-Decliner-Mails wirkt er als No-op.
- **Configure run after** → `Assistant_Forward_Mailto` → `Succeeded`.
- Rename → `Create_Reminder_Queue_Item`.

## Ablauf-Diagramm

```
Trigger: When a new email arrives (V3) — no_reply.events@deloitte.de / Inbox
   │
   ├── Is_Decline_Mail? (Subject startsWith Declined/Abgelehnt/Refusé/Rifiutato)
   │
   ├── Yes: Cleaned_Subject (strip prefix)
   │        │
   │        ├── Get_DEX_Event (Title eq Cleaned_Subject, mit Apostroph-Escape)
   │        │
   │        ├── Event_Found? (length > 0)
   │        │
   │        └── Yes: Get_Teilnehmer_Entry (ParticipantEmail = Sender, Status != Abgemeldet)
   │                   │
   │                   ├── Still_Registered? (length > 0)
   │                   │
   │                   └── Yes: Get_Existing_Reminder (Dedup check)
   │                            │
   │                            ├── No_Reminder_Yet? (length == 0)
   │                            │
   │                            └── Yes: Create_Reminder_Queue_Item in DEX_Emails
   │                                     (EmailType='OutlookDeclineReminder')
   │
   └── DEX_SEND_MAIL picks up Pending Mails und sendet mit Template,
       {{CancelUrl}} zeigt auf DEX.aspx?action=cancel&event=<eventNumber>
```

## App-seitige Unterstützung (bereits implementiert)

- `DexEventPlatform.tsx` parsed `?action=cancel&event=<n>` aus `window.location.search`
- Erkennt den Deep-Link schon beim ersten Render und zeigt einen Vollbild-
  Lade-Spinner (statt der LandingPage), solange die Events noch geladen werden.
  Dadurch sieht der User sofort, dass eine Aktion läuft und "hängt" nicht
  sekundenlang auf der Willkommensseite.
- Navigiert zu `my-events` mit `selectedEventId` + Intent `auto-cancel`
- `MyEventsPage.tsx` prüft den Intent, scrollt zur Event-Karte und **storniert
  die Registrierung direkt** (ohne zusätzlichen "Abmeldung bestätigen"-Klick).
  Der Klick auf den Action-Button in der Mail gilt als Bestätigung. Das ist
  sicher, weil der User eingeloggt sein muss und durch Item-Level Security
  ohnehin nur seine eigene Registrierung stornieren kann.
- Der OutlookDeclineReminder-Template-Body wird als fertig gewrapptes
  Deloitte-HTML (Logo, grüne Linie, Footer) in `DEX_EmailTemplates.BodyHtml`
  gespeichert — der `DEX_SEND_MAIL`-Flow ersetzt nur `{{LOGO_URL}}` und
  `{{ORB_URL}}` und muss keinen Template-Wrapper selbst erzeugen.
- Der Body endet schlicht mit "If you no longer want to attend, please also
  cancel your registration." / "Falls du nicht mehr teilnehmen möchtest,
  melde dich bitte auch offiziell ab." — ohne Waitlist-Bezug, damit die Mail
  auch für Events ohne Warteliste korrekt wirkt.

## Neues Template in DEX_EmailTemplates

Beim nächsten App-Start legt die App folgende Template-Einträge automatisch an (falls noch nicht vorhanden):

| TemplateType | Language | Subject |
|--------------|----------|---------|
| OutlookDeclineReminder | EN | Action Required: Do you also want to cancel your registration? {{EventTitle}} |
| OutlookDeclineReminder | DE | Action Required: Möchtest du dich auch offiziell abmelden? {{EventTitle}} |

Der Body enthält einen großen roten "Anmeldung stornieren" / "Cancel my registration"-Action-Button, der auf `{{CancelUrl}}` zeigt.

## Änderung in DEX_SEND_MAIL (bei der Template-Auswahl)

- Wenn `EmailType == 'OutlookDeclineReminder'`:
  - Template aus `DEX_EmailTemplates` mit `TemplateType eq 'OutlookDeclineReminder'` und passender Sprache holen
  - `{{CancelUrl}}` ersetzen mit Expression:
    ```
    concat(
      'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView&action=cancel&event=',
      first(body('Get_Event_from_DEX_Events')?['value'])?['EventNumber']
    )
    ```
    (EventNumber, nicht ID!)

## Finaler Flow-JSON (Stand 2026-04-14, Recipient-Filter via Filter array)

**Wichtig:** Die `Recipient`-Spalte in `DEX_Emails` ist ein **Note-Feld** (Multi-line text, weil es auch `;`-separierte Mehrfach-Empfaenger enthalten kann). SP-OData `$filter` erlaubt keinen `eq` auf Note-Feldern. Deshalb filtert `Get_Existing_Reminder` nur nach `EmailType + EventId`, und ein nachgeschalteter `Filter array`-Schritt (`Filter_By_Recipient`) pickt den aktuellen Sender heraus — mit Semikolon-Wrapping um Teil-Matches (z.B. `alice@x.de` in `alicebackup@x.de`) zu vermeiden.

**TRIGGER (Office 365 Outlook — SharedMailboxOnNewEmailV2):**
```json
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "mailboxAddress": "no_reply.events@deloitte.de",
      "importance": "Any",
      "hasAttachments": false,
      "includeAttachments": false,
      "folderId": "Inbox"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
      "connection": "shared_office365",
      "operationId": "SharedMailboxOnNewEmailV2"
    }
  },
  "recurrence": { "frequency": "Minute", "interval": 1 },
  "splitOn": "@triggerOutputs()?['body/value']"
}
```

**ROOT Condition `Is_Decline_Mail` + alle Kinder (If)** — Stand mit
Forward-Support (`Real_Sender`) und On-Behalf-Of-Support (`Decliner_Lastname`
/ `Decliner_Firstname` / `Final_Recipient_Email`):
```json
{
  "type": "If",
  "expression": {
    "and": [
      {
        "equals": [
          "@or(startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'declined:'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'declined '), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'abgelehnt:'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'abgelehnt '), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'refusé'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'rifiutato:'), and(or(startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'fw:'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'wg:')), or(contains(toLower(coalesce(triggerOutputs()?['body/bodyPreview'], '')), 'declined:'), contains(toLower(coalesce(triggerOutputs()?['body/bodyPreview'], '')), 'abgelehnt:'))))",
          "@true"
        ]
      }
    ]
  },
  "actions": {
    "Cleaned_Subject": {
      "type": "Compose",
      "inputs": "@trim(replace(replace(replace(replace(replace(replace(replace(replace(triggerOutputs()?['body/subject'], 'FW:', ''), 'WG:', ''), 'Declined:', ''), 'Declined ', ''), 'Abgelehnt:', ''), 'Abgelehnt ', ''), 'Refusé :', ''), 'Rifiutato:', ''))"
    },
    "Real_Sender": {
      "type": "Compose",
      "inputs": "@if(or(startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'fw:'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'wg:')), trim(first(split(first(skip(split(coalesce(triggerOutputs()?['body/bodyPreview'], ''), '<'), 1)), '>'))), triggerOutputs()?['body/from'])",
      "runAfter": { "Cleaned_Subject": ["Succeeded"] }
    },
    "Decliner_Lastname": {
      "type": "Compose",
      "inputs": "@if(contains(coalesce(triggerOutputs()?['body/bodyPreview'], ''), 'On Behalf Of '), trim(first(split(last(split(triggerOutputs()?['body/bodyPreview'], 'On Behalf Of ')), ','))), '')",
      "runAfter": { "Real_Sender": ["Succeeded"] }
    },
    "Decliner_Firstname": {
      "type": "Compose",
      "inputs": "@if(contains(coalesce(triggerOutputs()?['body/bodyPreview'], ''), 'On Behalf Of '), trim(replace(replace(replace(last(split(first(split(last(split(triggerOutputs()?['body/bodyPreview'], 'On Behalf Of ')), decodeUriComponent('%0D%0A'))), ',')), 'Prof. ', ''), 'Dr. ', ''), 'Dipl.-Ing. ', '')), '')",
      "runAfter": { "Decliner_Lastname": ["Succeeded"] }
    },
    "Get_DEX_Event": {
      "type": "OpenApiConnection",
      "inputs": {
        "parameters": {
          "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
          "table": "28457815-1163-4e92-8b08-3ae43f477d9e",
          "$filter": "@concat('Title eq ''', replace(outputs('Cleaned_Subject'), '''', ''''''), '''')",
          "$top": 1
        },
        "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "GetItems" }
      },
      "runAfter": { "Decliner_Firstname": ["Succeeded"] }
    },
    "Event_Found": {
      "type": "If",
      "expression": { "and": [ { "greater": ["@length(outputs('Get_DEX_Event')?['body/value'])", 0] } ] },
      "actions": {
        "Get_Teilnehmer_Entry": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "dataset": "@first(outputs('Get_DEX_Event')?['body/value'])?['SubsiteUrl']",
              "parameters/method": "GET",
              "parameters/uri": "@concat('_api/web/lists/getbytitle(''Teilnehmer'')/items?$filter=', if(empty(outputs('Decliner_Lastname')), concat('ParticipantEmail eq ''', replace(outputs('Real_Sender'), '''', ''''''), ''''), if(empty(outputs('Decliner_Firstname')), concat('Nachname eq ''', replace(outputs('Decliner_Lastname'), '''', ''''''), ''''), concat('Vorname eq ''', replace(outputs('Decliner_Firstname'), '''', ''''''), ''' and Nachname eq ''', replace(outputs('Decliner_Lastname'), '''', ''''''), ''''))), ' and Status ne ''Abgemeldet''&$top=1&$select=Id,Status,Vorname,Nachname,ParticipantEmail')",
              "parameters/headers": { "Accept": "application/json;odata=nometadata" }
            },
            "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "HttpRequest" }
          }
        },
        "Final_Recipient_Email": {
          "type": "Compose",
          "inputs": "@coalesce(first(body('Get_Teilnehmer_Entry')?['value'])?['ParticipantEmail'], outputs('Real_Sender'))",
          "runAfter": { "Get_Teilnehmer_Entry": ["Succeeded"] }
        },
        "Still_Registered": {
          "type": "If",
          "expression": { "and": [ { "greater": ["@length(body('Get_Teilnehmer_Entry')?['value'])", 0] } ] },
          "actions": {
            "Get_Existing_Reminder": {
              "type": "OpenApiConnection",
              "inputs": {
                "parameters": {
                  "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                  "table": "57aa0840-df98-41ae-a39b-323c0b80ae3b",
                  "$filter": "@concat('EmailType eq ''OutlookDeclineReminder'' and EventId eq ''', first(outputs('Get_DEX_Event')?['body/value'])?['ID'], '''')",
                  "$top": 20
                },
                "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "GetItems" }
              }
            },
            "Filter_By_Recipient": {
              "type": "Query",
              "inputs": {
                "from": "@body('Get_Existing_Reminder')?['value']",
                "where": "@contains(concat(';', replace(item()?['Recipient'], ' ', ''), ';'), concat(';', outputs('Final_Recipient_Email'), ';'))"
              },
              "runAfter": { "Get_Existing_Reminder": ["Succeeded"] }
            },
            "No_Reminder_Yet": {
              "type": "If",
              "expression": { "and": [ { "equals": ["@length(body('Filter_By_Recipient'))", 0] } ] },
              "actions": {
                "Get_Reminder_Template": {
                  "type": "OpenApiConnection",
                  "inputs": {
                    "parameters": {
                      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                      "table": "2c428d35-e6fb-42f9-8a20-580acd6d05f4",
                      "$filter": "@concat('TemplateType eq ''', if(empty(outputs('Decliner_Lastname')), 'OutlookDeclineReminder', 'OutlookDeclineReminder_OnBehalfOf'), ''' and Language eq ''', coalesce(first(outputs('Get_DEX_Event')?['body/value'])?['EmailLanguage'], 'EN'), '''')",
                      "$top": 1
                    },
                    "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "GetItems" }
                  }
                },
                "Assistant_Forward_Mailto": {
                  "type": "Compose",
                  "inputs": "@concat('mailto:', first(outputs('Get_DEX_Event')?['body/value'])?['OrganizerEmail'], '?subject=', encodeUriComponent(concat('Bitte um Abmeldung: ', first(body('Get_Teilnehmer_Entry')?['value'])?['Vorname'], ' ', first(body('Get_Teilnehmer_Entry')?['value'])?['Nachname'], ' \u2014 ', first(outputs('Get_DEX_Event')?['body/value'])?['Title'])), '&body=', encodeUriComponent(concat('Hallo,', decodeUriComponent('%0D%0A%0D%0A'), 'ich (Assistenz) habe in Vertretung den Outlook-Termin f\u00FCr ', first(body('Get_Teilnehmer_Entry')?['value'])?['Vorname'], ' ', first(body('Get_Teilnehmer_Entry')?['value'])?['Nachname'], ' (', outputs('Final_Recipient_Email'), ') abgelehnt.', decodeUriComponent('%0D%0A%0D%0A'), 'Bitte storniere die Anmeldung f\u00FCr die Person f\u00FCr das Event \"', first(outputs('Get_DEX_Event')?['body/value'])?['Title'], '\" \u00FCber das Admin Center der Event Experience Platform.', decodeUriComponent('%0D%0A%0D%0A'), 'Danke!')))",
                  "runAfter": { "Get_Reminder_Template": ["Succeeded"] }
                },
                "Create_Reminder_Queue_Item": {
                  "type": "OpenApiConnection",
                  "inputs": {
                    "parameters": {
                      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                      "table": "57aa0840-df98-41ae-a39b-323c0b80ae3b",
                      "item/Title": "@replace(coalesce(first(outputs('Get_Reminder_Template')?['body/value'])?['Subject'], concat('Outlook-Abmeldung-Reminder: ', first(outputs('Get_DEX_Event')?['body/value'])?['Title'])), '{{EventTitle}}', first(outputs('Get_DEX_Event')?['body/value'])?['Title'])",
                      "item/Recipient": "@outputs('Final_Recipient_Email')",
                      "item/RecipientName": "@coalesce(first(body('Get_Teilnehmer_Entry')?['value'])?['Vorname'], outputs('Final_Recipient_Email'))",
                      "item/EmailType/Value": "OutlookDeclineReminder",
                      "item/EventTitle": "@first(outputs('Get_DEX_Event')?['body/value'])?['Title']",
                      "item/Status/Value": "Pending",
                      "item/Body": "@replace(replace(replace(replace(coalesce(first(outputs('Get_Reminder_Template')?['body/value'])?['BodyHtml'], ''), '{{Name}}', coalesce(first(body('Get_Teilnehmer_Entry')?['value'])?['Vorname'], outputs('Final_Recipient_Email'))), '{{EventTitle}}', first(outputs('Get_DEX_Event')?['body/value'])?['Title']), '{{CancelUrl}}', concat('https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView&action=cancel&event=', string(first(outputs('Get_DEX_Event')?['body/value'])?['EventNumber']))), '{{AssistantForwardUrl}}', outputs('Assistant_Forward_Mailto'))",
                      "item/EventId": "@first(outputs('Get_DEX_Event')?['body/value'])?['ID']"
                    },
                    "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "PostItem" }
                  },
                  "runAfter": { "Assistant_Forward_Mailto": ["Succeeded"] }
                }
              },
              "else": { "actions": {} },
              "runAfter": { "Filter_By_Recipient": ["Succeeded"] }
            }
          },
          "else": { "actions": {} },
          "runAfter": { "Final_Recipient_Email": ["Succeeded"] }
        }
      },
      "else": { "actions": {} },
      "runAfter": { "Get_DEX_Event": ["Succeeded"] }
    }
  },
  "else": { "actions": {} },
  "runAfter": {}
}
```

---

## Archiv: DEX_OutlookDeclineHandler v1 (event-modified, deprecated 2026-04-14)

Die erste Version nutzte den `When an event is modified (V3)`-Trigger und filterte `body/attendees[*].status.response == 'declined'`. In der Praxis gibt der Shared-Mailbox-Trigger dieses pro-Attendee-Feld jedoch **nicht** zuverlässig zurück (nur `requiredAttendees`/`optionalAttendees` als Semicolon-String). Der Flow hat daher nie eine Reminder-Mail erzeugt. Die Mail-basierte Variante oben ersetzt diesen Ansatz.

Der v1-Flow ist deaktiviert aber nicht gelöscht (`DEX_OutlookDeclineHandler_v1_old`), kann nach 1 Woche stabiler Mail-Variante endgültig entfernt werden.

---

## 6. DEX_OutlookForwardHandler

**Trigger:** Neue Mail in `no_reply.events@deloitte.de` mit Subject startend mit `Meeting Forward Notification:` (EN) oder `Terminweiterleitungsbenachrichtigung:` (DE)
**Zweck:** Wenn ein Teilnehmer einen Outlook-Termin an Dritte weiterleitet, bekommt `no_reply.events@deloitte.de` eine Info-Mail. Der Flow prüft, ob die weitergeleitete Person bereits in der SharePoint-Teilnehmerliste eingetragen ist. Wenn **nein**, geht eine FYI-Mail an den Organizer raus (Template `OutlookForwardNotification` aus `DEX_EmailTemplates`, vom Flow gerendert, Body landet in `DEX_Emails`-Queue → DEX_SEND_MAIL versendet).
**Letztes Update:** 2026-04-15 (FW:/WG:-Varianten abgedeckt, Cleaned_Subject auf `last(split)` umgestellt)
**Listen-GUIDs:** DEX_Events `28457815-1163-4e92-8b08-3ae43f477d9e`, DEX_EmailTemplates `2c428d35-e6fb-42f9-8a20-580acd6d05f4`, DEX_Emails `57aa0840-df98-41ae-a39b-323c0b80ae3b`

### Hintergrund

Beispiel-Mail, die das auslöst:

```
From: Microsoft Outlook on behalf of von Rueden, Dr. Michael
Subject: Meeting Forward Notification: E2E M&A Activation Session
Body:
  Your meeting was forwarded
  von Rueden, Dr. Michael has forwarded your meeting request to additional recipients.

  Meeting:       E2E M&A Activation Session
  Meeting Time:  Thursday, 23 April 2026, 19:00 to Friday, 24 April 2026, 15:30.
  Recipients:    Mauß, Anna Kristina
```

Problem: Anna Kristina Mauß ist im Outlook-Termin drin — aber **nicht** in der SharePoint-Teilnehmerliste der Event-Subsite. Sie hat keine TeilnehmerID, keinen QR-Code, und der Organizer sieht sie nicht in der App. Dazu kommt: sie könnte ausserhalb der Audience-Filter (Location/JobTitle) liegen und eigentlich gar nicht teilnehmen dürfen.

### Ablauf

1. Trigger (neue Mail im Inbox von `no_reply.events@deloitte.de`)
2. **Is_ForwardNotification?** → Subject startet mit `Meeting Forward Notification:` ODER `Terminweiterleitungsbenachrichtigung:`
3. **Cleaned_Subject** → Event-Titel aus Subject extrahieren (alles nach dem `:`)
4. **Get_DEX_Event** → SharePoint Get items `DEX_Events` mit `Title eq '<Cleaned_Subject>'`
5. **Event_Found?** → weiter nur wenn Event existiert
6. **Parse_Recipients** → Recipient-Namen aus `body/body` extrahieren (alle Namen nach `Recipients` / `Empfänger` bis `All times listed`)
7. **Resolve_Recipient_Email** → Graph API User-Search per DisplayName (`Nachname, Vorname`) → Email
8. **Get_Teilnehmer_Entry** → SharePoint HTTP request auf Subsite-Teilnehmerliste: gibt es einen Eintrag mit `ParticipantEmail eq '<resolved_email>'` der nicht `Abgemeldet` ist?
9. **Already_Registered?** → wenn **ja**: Flow endet (Log-Eintrag "OK, schon eingeladen"). Wenn **nein**: weiter zu 10.
10. **Get_ForwardTemplate** → Template `OutlookForwardNotification` (DE/EN je nach Event) aus `DEX_EmailTemplates` laden. Das Template wird **von der App automatisch angelegt** (siehe Abschnitt "Template-Seeding durch die App") — inklusive Deloitte-Outlook-Wrapper.
11. **Rendered_Subject / Rendered_Body** → Platzhalter `{{EventTitle}}`, `{{Forwarder}}`, `{{Recipient}}`, `{{RecipientEmail}}`, `{{OrganizerFirstName}}`, `{{AppUrl}}` per `replace()` ersetzen.
12. **Create_FYI_Email** → DEX_Emails-Queue-Item anlegen mit den echten Spalten:
    - `Title` = Rendered Subject
    - `Recipient` = Organizer-Email (Plain-Text)
    - `Body` = Rendered HTML-Body (fertig gerendert, kein weiteres Template-Lookup im DEX_SEND_MAIL nötig)
    - `EmailType` = `Info`, `Status` = `Pending`, `EventId` = Event-ID, `EventTitle` = Cleaned_Subject

### SharePoint-Liste DEX_Outlook (unverändert)

Dieser Flow nutzt **nicht** die DEX_Outlook-Queue (da keine Outlook-Änderung getriggert wird — der Outlook-Termin hat den neuen Empfänger ja schon). Er nutzt die bestehende `DEX_Emails`-Queue für die FYI-Mail.

### UI-Anleitung zum Anlegen (Schritt für Schritt)

**1. Neuer Cloud-Flow anlegen**
1. https://make.powerautomate.com öffnen
2. Links **+ Create** → **Automated cloud flow**
3. **Flow name:** `DEX_OutlookForwardHandler`
4. Trigger: `When a new email arrives V3` → **Office 365 Outlook — When a new email arrives (V3)** auswählen
5. **Create**

**2. Trigger konfigurieren**
- **Folder:** `Inbox`
- **Show advanced options:**
  - **Original Mailbox Address:** `no_reply.events@deloitte.de`
  - **Include Attachments:** `No`
  - **Subject Filter:** leer lassen (Language-Filter via Condition in Schritt 3)
- Settings → Concurrency: **1** (sequentiell, maximumWaitingRuns 100)

**3. Condition `Is_ForwardNotification`**
- Linke Seite: **Expression (fx)** →
  ```
  or(
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'meeting forward notification:'),
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'terminweiterleitungsbenachrichtigung:')
  )
  ```
- Operator: `is equal to`
- Rechte Seite: **Expression (fx)** → `true`
- Wenn **no** → Terminate (Succeeded). Wenn **yes** → weiter.

**4. `Cleaned_Subject` (Compose)**
- Expression:
  ```
  trim(
    last(
      split(triggerOutputs()?['body/subject'], ':')
    )
  )
  ```
- Achtung: `last(split(...,':'))` kann bei Event-Titeln mit Doppelpunkt falsch sein. Alternativ robuster:
  ```
  trim(
    substring(
      triggerOutputs()?['body/subject'],
      add(indexOf(triggerOutputs()?['body/subject'], ':'), 1)
    )
  )
  ```
- Rename → `Cleaned_Subject`

**5. `Get_DEX_Event` (SharePoint Get items)**
- Site Address: `https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform`
- List Name: `DEX_Events`
- Filter Query: **Expression (fx)** →
  ```
  concat('Title eq ''', replace(outputs('Cleaned_Subject'), '''', ''''''), '''')
  ```
- Top Count: `1`

**6. Condition `Event_Found`**
- Expression: `length(outputs('Get_DEX_Event')?['body/value'])` `is greater than` `0`
- Wenn **no** → Terminate (Succeeded, Message `Event not found in DEX_Events`)
- Wenn **yes** → weiter

**7. `Parse_Recipients` (Compose)**

Die Mail enthält die Recipient-Namen im HTML-Body. Der `bodyPreview` (Plaintext) hat das Format:
```
... Recipients   Mauß, Anna Kristina   All times listed ...
```
Expression zum Extrahieren:
```
trim(
  first(
    split(
      last(
        split(
          coalesce(triggerOutputs()?['body/bodyPreview'], ''),
          'Recipients'
        )
      ),
      'All times listed'
    )
  )
)
```
Für DE-Mails zusätzlich auch `Empfänger` als Split-Token. Alternative (robuster):
```
if(
  contains(coalesce(triggerOutputs()?['body/bodyPreview'], ''), 'Recipients'),
  trim(first(split(last(split(triggerOutputs()?['body/bodyPreview'], 'Recipients')), 'All times listed'))),
  trim(first(split(last(split(triggerOutputs()?['body/bodyPreview'], 'Empfänger')), 'Alle aufgeführten Zeiten')))
)
```
Rename → `Recipient_DisplayName`.

**8. `Resolve_Recipient_Email` (Office 365 Users — Search for users (V2))**
- Search Term: `@outputs('Recipient_DisplayName')`
- Top: `5`
- Nach der Action: `Filter_Matching_User` (Data Operation — Filter array) über `body('Resolve_Recipient_Email')?['value']` mit `item()?['displayName'] is equal to @outputs('Recipient_DisplayName')` (exakte Match).
- Compose `Recipient_Email`: `first(body('Filter_Matching_User'))?['mail']`

**9. Condition `Email_Resolved`**
- Expression: `outputs('Recipient_Email')` `is not equal to` `null`
- Wenn **no** → FYI-Mail an Organizer mit Hinweis "Recipient-Email konnte nicht aufgelöst werden: <DisplayName>" (siehe Schritt 12, aber mit anderer Nachricht)
- Wenn **yes** → weiter

**10. `Get_Teilnehmer_Entry` (Send HTTP request to SharePoint)**

Die Teilnehmer-Liste liegt auf der Event-Subsite. `SubsiteUrl` aus `Get_DEX_Event` holen:
- Site Address: `@{first(outputs('Get_DEX_Event')?['body/value'])?['SubsiteUrl']}`
- Method: `GET`
- Uri:
  ```
  _api/web/lists/getbytitle('Teilnehmer')/items?$filter=ParticipantEmail eq '@{outputs('Recipient_Email')}' and Status ne 'Abgemeldet'&$top=1
  ```
- Headers: `Accept: application/json;odata=nometadata`

**11. Condition `Already_Registered`**
- Expression: `length(body('Get_Teilnehmer_Entry')?['value'])` `is greater than` `0`
- Wenn **yes** → Terminate (Succeeded, Message `Recipient already registered`)
- Wenn **no** → weiter zu 12

**12. `Get_ForwardTemplate` (SharePoint Get items auf DEX_EmailTemplates)**

Das Template `OutlookForwardNotification` wird **von der App automatisch angelegt** (siehe unten, Abschnitt "Template-Seeding durch die App"). Der Flow holt es per Filter. Sprache wird vom Event übernommen (`EmailLanguage`-Feld), mit Fallback auf `EN`.

- Site Address: `https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform`
- List Name: `DEX_EmailTemplates`
- Filter Query: **Expression (fx)** →
  ```
  concat('TemplateType eq ''OutlookForwardNotification'' and Language eq ''', coalesce(first(outputs('Get_DEX_Event')?['body/value'])?['EmailLanguage'], 'EN'), '''')
  ```
- Top Count: `1`

**12a. `Rendered_Subject` (Compose)**

Platzhalter im Subject ersetzen. Expression (fx):
```
replace(replace(replace(replace(replace(first(outputs('Get_ForwardTemplate')?['body/value'])?['Subject'], '{{EventTitle}}', outputs('Cleaned_Subject')), '{{Forwarder}}', triggerOutputs()?['body/from']), '{{Recipient}}', outputs('Recipient_DisplayName')), '{{RecipientEmail}}', coalesce(outputs('Recipient_Email'), 'nicht aufgelöst')), '{{OrganizerFirstName}}', first(split(first(split(first(outputs('Get_DEX_Event')?['body/value'])?['Organizer'], ';')), ' ')))
```

**12b. `Rendered_Body` (Compose)**

Die gleiche Replace-Kaskade auf dem `BodyHtml`-Feld, zusätzlich noch `{{AppUrl}}`. Expression (fx):
```
replace(replace(replace(replace(replace(replace(first(outputs('Get_ForwardTemplate')?['body/value'])?['BodyHtml'], '{{EventTitle}}', outputs('Cleaned_Subject')), '{{Forwarder}}', triggerOutputs()?['body/from']), '{{Recipient}}', outputs('Recipient_DisplayName')), '{{RecipientEmail}}', coalesce(outputs('Recipient_Email'), 'nicht aufgelöst')), '{{OrganizerFirstName}}', first(split(first(split(first(outputs('Get_DEX_Event')?['body/value'])?['Organizer'], ';')), ' '))), '{{AppUrl}}', 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView')
```

**13. `Create_FYI_Email` (SharePoint Create item, Liste DEX_Emails)**

Schreibt den fertigen Mail-Body in die Queue. `DEX_SEND_MAIL` verschickt ihn, ohne selbst Templates laden zu müssen — der komplette, schon gerenderte HTML-Body steht direkt im `Body`-Feld.

Echte Spalten von `DEX_Emails` (siehe `EventService.ts:196`):

| Feld | Typ | Wert für diesen Flow |
|------|-----|----------------------|
| `Title` | Text | = Subject (`outputs('Rendered_Subject')`) |
| `Recipient` | Note, Plain-Text | Organizer-Email (erste aus `;`-Liste) |
| `Cc` | Note, Plain-Text | leer |
| `RecipientName` | Text | = `first(split(first(outputs('Get_DEX_Event')?['body/value'])?['Organizer'], ';'))` |
| `Body` | Note (HTML) | = `outputs('Rendered_Body')` |
| `EmailType` | Choice (Anmeldung, Abmeldung, Warteliste, Nachruecken, **Info**) | `Info` |
| `EventTitle` | Text | = `outputs('Cleaned_Subject')` |
| `EventId` | Text | = `string(first(outputs('Get_DEX_Event')?['body/value'])?['Id'])` |
| `Status` | Choice (**Pending**, Sent, Failed) | `Pending` |
| `SentDate` | Date | leer (Flow setzt es beim Versand) |

Der `DEX_SEND_MAIL`-Flow pickt das Item auf (`Status eq 'Pending'`) und verschickt es per Office-365-Mail-Action — **keine Anpassung an DEX_SEND_MAIL nötig**, weil der Body schon fertig gerendert ist.

### Template-Seeding durch die App

Das Template `OutlookForwardNotification` (DE + EN) wird **automatisch beim App-Start** in die Liste `DEX_EmailTemplates` geschrieben. Zuständig sind drei Funktionen in `EventService.ts`:

- `ensureEmailTemplatesList()` — erstmaliges Anlegen der Liste mit allen Default-Templates
- `ensureMissingEmailTemplates()` — Nachlegen auf Tenants, wo die Liste schon existiert, aber `OutlookForwardNotification` noch fehlt
- `upgradeStandardEmailTemplates()` — Updated den `BodyHtml` wenn das Template in der App-Version verändert wurde (überschreibt User-Customizing)

Der HTML-Body wird über `wrapTemplateForStorage()` aus `services/EmailTemplates.ts` gerendert — damit sieht die FYI-Mail exakt wie die anderen DEX-Mails aus (schwarzer Deloitte-Header mit Logo, grüne Trennlinie, DEX-Orb, Content-Block, Footer mit Legal-Text).

**Platzhalter im Body (werden vom Flow per `replace()` ersetzt, siehe 12a/12b):**
- `{{OrganizerFirstName}}` — Vorname des Organizers
- `{{Forwarder}}` — Name/Email der Person, die den Termin weitergeleitet hat (aus `body/from`)
- `{{Recipient}}` — Name der hinzugefügten Person
- `{{RecipientEmail}}` — Resolved Email oder `nicht aufgelöst`
- `{{EventTitle}}` — Event-Titel
- `{{AppUrl}}` — DEX-App-URL (hardcoded im Flow)

**Code-Konstanten:** `OUTLOOK_FORWARD_BODY_DE` / `OUTLOOK_FORWARD_BODY_EN` in `EventService.ts`.

### Änderung in DEX_SEND_MAIL

**Keine Änderung nötig.** Der Flow verschickt einfach das fertige `Body`-Feld aus DEX_Emails per Office-365-Mail-Action — das neue Template wird für DEX_SEND_MAIL transparent behandelt (wie alle `Info`-Mails).

### Bekannte Einschränkungen

| Fall | Verhalten |
|------|-----------|
| Forwarder hat Outlook auf DE → `Terminweiterleitungsbenachrichtigung:` | ✅ (Condition `Is_ForwardNotification` deckt beide Sprachen ab) |
| Forwarder hat Outlook auf EN → `Meeting Forward Notification:` | ✅ |
| Notification wurde nochmal weitergeleitet (`FW:` / `WG:` davor) | ✅ Seit 2026-04-15: Condition kennt alle 6 Varianten (DE/EN × direkt/FW:/WG:). `body/from` ist dann allerdings der Weiterleiter, nicht der originale Forwarder — `{{Forwarder}}` in der FYI-Mail kann deshalb ungenau sein. |
| Forwarder hat andere Sprache (FR/IT/…) | ❌ Subject wird nicht erkannt — erweitern bei Bedarf |
| Recipient ist ein externer User (kein Azure AD Account) | ❌ `Resolve_Recipient_Email` findet keinen Treffer → `Email_Resolved`-else-Zweig ist leer, Flow terminiert ohne Mail |
| Recipient-Name steht in uneindeutiger Form (nur Vorname, Firmenkürzel) | ⚠️ Graph-Search kann 0 oder mehrere Treffer liefern → `Filter_Matching_User` per exakter `displayName`-Gleichheit |
| Mehrere Recipients in einer Forward-Notification | ⚠️ Flow behandelt nur den ersten — für Multi-Recipient Schleife über gesplittete Namen nötig |
| Event-Titel enthält Doppelpunkt | ⚠️ `Cleaned_Subject` per `last(split(..., ':'))` nimmt nur das letzte Segment → bei Events wie `DEX: Sommer-Event` landet nur `Sommer-Event` im Cleaned_Subject und `Get_DEX_Event` findet nichts |

### Teststrategie

1. Manuell in Outlook einen Testtermin aus `no_reply.events@deloitte.de` erstellen, einladen.
2. Als eingeladener User: **Forward** auf eine andere Person (Deloitte-interne) → Flow sollte die FYI-Mail an den Organizer queuen.
3. Als eingeladener User: **Forward** auf eine Person, die bereits registriert ist → Flow sollte **keine** FYI schicken.
4. Als eingeladener User: **Forward** auf einen Externen (keine Azure AD-Identität) → Email_Resolved-Condition sollte greifen, Flow terminiert sauber (aktuell keine FYI bei nicht-auflösbarer Mail — Verbesserungspotenzial).
5. Flow-Runs kontrollieren: jeder Schritt sollte `Succeeded` sein, Terminate-Zweige dokumentieren warum keine Mail verschickt wurde.

### Finaler Flow-JSON (Stand 2026-04-15, HTML-Parser für Name + Email, ohne Graph-Search)

**Key Changes vs. initiale Version:**
- `Recipient_DisplayName` extrahiert jetzt den Namen **direkt aus dem HTML-Body** (via `mailto:...">Name</a>`-Muster), nicht mehr über Graph-API-Search
- `Recipient_Email` ebenfalls direkt aus dem HTML — keine Office-365-User-Search mehr nötig
- Dadurch: `Resolve_Recipient_Email` (Office 365 Users) und `Filter_Matching_User` (Filter array) **entfernt** — Flow ist schneller und robuster bei Externen
- `Cleaned_Subject` auf `last(split(subject, ':'))` umgestellt (robust gegen FW:/WG:-Prefixes)
- `Get_DEX_Event` auf `substringof(subject, Title)` statt `eq` (matcht Events mit längerem Titel wie `E2E M&A Activation Session Munich`)
- `Email_Resolved` nutzt `empty()` statt `null`-Vergleich
- `Rendered_Subject`/`Rendered_Body`: 3 Platzhalter-Fixes (OrganizerFirstName aus `Nachname, Vorname`, Forwarder aus `bodyPreview`, Recipient-Fallback auf Email)



TRIGGER:
```json
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": { "mailboxAddress": "no_reply.events@deloitte.de" },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
      "connection": "shared_office365",
      "operationId": "SharedMailboxOnNewEmailV2"
    }
  },
  "recurrence": { "frequency": "Minute", "interval": 1 },
  "runtimeConfiguration": { "concurrency": { "runs": 1, "maximumWaitingRuns": 100 } },
  "splitOn": "@triggerOutputs()?['body/value']"
}
```

IS_FORWARDNOTIFICATION (If):
```json
{
  "type": "If",
  "expression": {
    "and": [
      {
        "equals": [
          "@or(or(startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'meeting forward notification:'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'terminweiterleitungsbenachrichtigung:')), or(or(startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'fw: meeting forward notification:'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'wg: meeting forward notification:')), or(startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'fw: terminweiterleitungsbenachrichtigung:'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'wg: terminweiterleitungsbenachrichtigung:'))))",
          true
        ]
      }
    ]
  },
  "actions": {
    "Cleaned_Subject": {
      "type": "Compose",
      "inputs": "@trim(last(split(coalesce(triggerOutputs()?['body/subject'], ''), ':')))"
    },
    "Get_DEX_Event": {
      "type": "OpenApiConnection",
      "inputs": {
        "parameters": {
          "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
          "table": "28457815-1163-4e92-8b08-3ae43f477d9e",
          "$filter": "@concat('Title eq ''', replace(outputs('Cleaned_Subject'), '''', ''''''), '''')",
          "$top": 1
        },
        "host": {
          "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
          "connection": "shared_sharepointonline",
          "operationId": "GetItems"
        }
      },
      "runAfter": { "Cleaned_Subject": ["Succeeded"] }
    },
    "Event_Found": {
      "type": "If",
      "expression": {
        "and": [ { "greater": ["@length(outputs('Get_DEX_Event')?['body/value'])", 0] } ]
      },
      "actions": {
        "Recipient_DisplayName": {
          "type": "Compose",
          "inputs": "@if(contains(coalesce(triggerOutputs()?['body/bodyPreview'], ''), 'Recipients'), trim(first(split(last(split(triggerOutputs()?['body/bodyPreview'], 'Recipients')), 'All times listed'))), trim(first(split(last(split(triggerOutputs()?['body/bodyPreview'], 'Empfänger')), 'Alle aufgeführten Zeiten'))))"
        },
        "Resolve_Recipient_Email": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "top": 5,
              "isSearchTermRequired": true,
              "searchTerm": "@outputs('Recipient_DisplayName')"
            },
            "host": {
              "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365users",
              "connection": "shared_office365users",
              "operationId": "SearchUserV2"
            }
          },
          "runAfter": { "Recipient_DisplayName": ["Succeeded"] }
        },
        "Filter_Matching_User": {
          "type": "Query",
          "inputs": {
            "from": "@body('Resolve_Recipient_Email')?['value']",
            "where": "@equals(item()?['displayName'],outputs('Recipient_DisplayName'))"
          },
          "runAfter": { "Resolve_Recipient_Email": ["Succeeded"] }
        },
        "Recipient_Email": {
          "type": "Compose",
          "inputs": "@first(body('Filter_Matching_User'))?['mail']",
          "runAfter": { "Filter_Matching_User": ["Succeeded"] }
        },
        "Email_Resolved": {
          "type": "If",
          "expression": {
            "and": [ { "equals": ["@empty(outputs('Recipient_Email'))", false] } ]
          },
          "actions": {
            "Get_Teilnehmer_Entry": {
              "type": "OpenApiConnection",
              "inputs": {
                "parameters": {
                  "dataset": "@first(outputs('Get_DEX_Event')?['body/value'])?['SubsiteUrl']",
                  "parameters/method": "GET",
                  "parameters/uri": "_api/web/lists/getbytitle('Teilnehmer')/items?$filter=ParticipantEmail eq '@{outputs('Recipient_Email')}' and Status ne 'Abgemeldet'&$top=1",
                  "parameters/headers": { "Accept": "application/json;odata=nometadata" }
                },
                "host": {
                  "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                  "connection": "shared_sharepointonline",
                  "operationId": "HttpRequest"
                }
              }
            },
            "Already_Registered": {
              "type": "If",
              "expression": {
                "and": [ { "greater": ["@length(body('Get_Teilnehmer_Entry')?['value'])", 0] } ]
              },
              "actions": {
                "Terminate_2": { "type": "Terminate", "inputs": { "runStatus": "Succeeded" } }
              },
              "else": {
                "actions": {
                  "Get_ForwardTemplate": {
                    "type": "OpenApiConnection",
                    "inputs": {
                      "parameters": {
                        "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                        "table": "2c428d35-e6fb-42f9-8a20-580acd6d05f4",
                        "$filter": "@concat('TemplateType eq ''OutlookForwardNotification'' and Language eq ''', coalesce(first(outputs('Get_DEX_Event')?['body/value'])?['EmailLanguage'], 'EN'), '''')",
                        "$top": 1
                      },
                      "host": {
                        "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                        "connection": "shared_sharepointonline",
                        "operationId": "GetItems"
                      }
                    }
                  },
                  "Rendered_Subject": {
                    "type": "Compose",
                    "inputs": "@replace(replace(replace(replace(replace(first(outputs('Get_ForwardTemplate')?['body/value'])?['Subject'], '{{EventTitle}}', outputs('Cleaned_Subject')), '{{Forwarder}}', triggerOutputs()?['body/from']), '{{Recipient}}', outputs('Recipient_DisplayName')), '{{RecipientEmail}}', coalesce(outputs('Recipient_Email'), 'nicht aufgelöst')), '{{OrganizerFirstName}}', first(split(first(split(first(outputs('Get_DEX_Event')?['body/value'])?['Organizer'], ';')), ' ')))",
                    "runAfter": { "Get_ForwardTemplate": ["Succeeded"] }
                  },
                  "Rendered_Body": {
                    "type": "Compose",
                    "inputs": "@replace(replace(replace(replace(replace(replace(first(outputs('Get_ForwardTemplate')?['body/value'])?['BodyHtml'], '{{EventTitle}}', outputs('Cleaned_Subject')), '{{Forwarder}}', triggerOutputs()?['body/from']), '{{Recipient}}', outputs('Recipient_DisplayName')), '{{RecipientEmail}}', coalesce(outputs('Recipient_Email'), 'nicht aufgelöst')), '{{OrganizerFirstName}}', first(split(first(split(first(outputs('Get_DEX_Event')?['body/value'])?['Organizer'], ';')), ' '))), '{{AppUrl}}', 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView')",
                    "runAfter": { "Rendered_Subject": ["Succeeded"] }
                  },
                  "Create_FYI_Email": {
                    "type": "OpenApiConnection",
                    "inputs": {
                      "parameters": {
                        "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                        "table": "57aa0840-df98-41ae-a39b-323c0b80ae3b",
                        "item/Title": "@outputs('Rendered_Subject')",
                        "item/Recipient": "@first(split(first(outputs('Get_DEX_Event')?['body/value'])?['OrganizerEmail'], ';'))",
                        "item/RecipientName": "@first(split(first(outputs('Get_DEX_Event')?['body/value'])?['Organizer'], ';'))",
                        "item/Body": "@outputs('Rendered_Body')",
                        "item/EmailType/Value": "Info",
                        "item/EventTitle": "@outputs('Cleaned_Subject')",
                        "item/EventId": "@string(first(outputs('Get_DEX_Event')?['body/value'])?['Id'])",
                        "item/Status/Value": "Pending"
                      },
                      "host": {
                        "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                        "connection": "shared_sharepointonline",
                        "operationId": "PostItem"
                      }
                    },
                    "runAfter": { "Rendered_Body": ["Succeeded"] }
                  }
                }
              },
              "runAfter": { "Get_Teilnehmer_Entry": ["Succeeded"] }
            }
          },
          "else": { "actions": {} },
          "runAfter": { "Recipient_Email": ["Succeeded"] }
        }
      },
      "else": {
        "actions": { "Terminate_1": { "type": "Terminate", "inputs": { "runStatus": "Succeeded" } } }
      },
      "runAfter": { "Get_DEX_Event": ["Succeeded"] }
    }
  },
  "else": {
    "actions": { "Terminate": { "type": "Terminate", "inputs": { "runStatus": "Succeeded" } } }
  },
  "runAfter": {}
}
```

### Offenes Verbesserungspotenzial

- **`Email_Resolved`-else-Zweig ist leer:** Wenn die weitergeleitete Person nicht per Graph-Search aufgelöst werden kann (z.B. Externe, Gastaccount mit anderem DisplayName-Format), verliert der Flow sie stumm. Besser wäre ein zweiter `Create_FYI_Email`-Block in diesem Zweig, der eine FYI-Mail mit `RecipientEmail = "nicht aufgelöst — bitte manuell prüfen: <DisplayName>"` an den Organizer schickt.
- **Mehrere Recipients:** Die `bodyPreview`-Parsing-Logik nimmt aktuell nur den ersten Namen. Wenn jemand den Termin an mehrere Personen gleichzeitig forwarded (z.B. "Mauß, Anna Kristina; Müller, Max"), sollte eine Apply-to-each-Schleife über `split(outputs('Recipient_DisplayName'), ';')` iterieren.
- **Subject-Parsing bei `:` im Event-Titel:** `Cleaned_Subject` per `substring(..., indexOf(':'))` schneidet nach dem **ersten** `:`. Bei Events wie `DEX: Sommer-Event` landet nur `Sommer-Event` im Cleaned_Subject und `Get_DEX_Event` findet nichts. Robustere Variante: per `add(indexOf('notification:'), 13)` bzw. `add(indexOf('benachrichtigung:'), 17)`.
