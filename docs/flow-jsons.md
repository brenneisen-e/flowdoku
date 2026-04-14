# Power Automate Flow JSONs

Dieses Dokument enthält die vollständigen Flow-Definitionen aller 4 DEX-Flows.
Wird aktualisiert wenn Flows geändert werden.

---

## 1. DEX_IDReorder_TeilnehmerIDs

**Trigger:** Neuer Eintrag in DEX_IDReorder
**Zweck:** TeilnehmerIDs neu vergeben + Nachrücken von Warteliste
**Letztes Update:** 2026-04-06

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

GET_ACTIVE_PARTICIPANTS:
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "@outputs('Settings')?['siteAddress']",
      "parameters/method": "GET",
      "parameters/uri": "_api/web/lists/getbytitle('@{outputs('Settings')?['listName']}')/items?$select=Id,TeilnehmerID,Status,RegistrationDate&$filter=(Status eq 'Angemeldet') or (Status eq 'QR versendet') or (Status eq 'Eingecheckt')&$orderby=RegistrationDate asc&$top=5000",
      "parameters/headers": { "Accept": "application/json;odata=nometadata" }
    },
    "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "HttpRequest" }
  },
  "runAfter": { "Get_ListItemType": ["Succeeded"] }
}

GENERATE_INDICES:
{
  "type": "Compose",
  "inputs": "@range(0, length(body('Get_Active_Participants')?['value']))",
  "runAfter": { "Get_Active_Participants": ["Succeeded"] }
}

GENERATESPDATA:
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
                  "parameters/body": "{\"__metadata\":{\"type\":\"SP.Data.TeilnehmerListItem\"},\"Status\":\"Angemeldet\",\"TeilnehmerID\":@{add(length(body('GenerateSPData')), 1)}}"
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
                  "item/Title": "@replace(replace(coalesce(first(body('Get_Email_Template')?['d']?['results'])?['Subject'], concat('Platz frei: ', triggerOutputs()?['body/Title'])), '{{Name}}', first(body('Get_Waitlist_First')?['d']?['results'])?['ParticipantName']), '{{EventTitle}}', triggerOutputs()?['body/Title'])",
                  "item/Recipient": "@first(body('Get_Waitlist_First')?['d']?['results'])?['ParticipantEmail']",
                  "item/RecipientName": "@first(body('Get_Waitlist_First')?['d']?['results'])?['ParticipantName']",
                  "item/EmailType/Value": "Nachruecken",
                  "item/EventTitle": "@triggerOutputs()?['body/Title']",
                  "item/Status/Value": "Pending",
                  "item/Body": "@replace(replace(coalesce(first(body('Get_Email_Template')?['d']?['results'])?['BodyHtml'], concat('Hallo ', first(body('Get_Waitlist_First')?['d']?['results'])?['ParticipantName'], ', du bist nachgerückt!')), '{{Name}}', first(body('Get_Waitlist_First')?['d']?['results'])?['ParticipantName']), '{{EventTitle}}', triggerOutputs()?['body/Title'])",
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
                  "item/Title": "Einladen: @{triggerOutputs()?['body/Title']}",
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

**Trigger:** Neuer Eintrag in DEX_Outlook
**Zweck:** Outlook-Termin verwalten: Teilnehmer einladen/ausladen, Event-Daten aktualisieren, Termin loeschen (via Graph API)
**Letztes Update:** 2026-04-14 (DeleteEvent-Branch hinzugefuegt)

Ablauf:
1. Trigger (neues DEX_Outlook-Item)
2. **Is_DeleteEvent?** → Ja: Outlook-Termin per `triggerBody()?['CalendarLink']` finden (iCalUId) → DELETE → Status=Sent. Nein: weiter.
3. Event-Details laden (DEX_Events via EventId) → CalendarLink vorhanden? → Outlook-Event per iCalUId finden → Event-ID speichern → Event gefunden?
4. Bestehende Attendees laden → Is_UpdateEvent? → Ja: PATCH Titel/Start/Ende → Nein: Einladen/Ausladen → Status=Sent

**ActionTypes:**
- `Einladen` — einzelnen Teilnehmer zum Outlook-Termin hinzufuegen
- `Ausladen` — einzelnen Teilnehmer aus dem Outlook-Termin entfernen
- `UpdateEvent` — Titel/Start/Ende des Outlook-Termins aktualisieren
- `DeleteEvent` — kompletten Outlook-Termin loeschen. Das DEX_Outlook-Queue-Item
  enthaelt `CalendarLink` (iCalUId) direkt, weil das zugehoerige DEX_Events-Item
  bereits geloescht ist, wenn der Flow laeuft.

**Concurrency:** 1 (sequentielle Verarbeitung, max 100 wartende Runs)

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
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'rifiutato:')
  )
  ```
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
              replace(triggerOutputs()?['body/subject'], 'Declined:', ''),
              'Declined ', ''),
            'Abgelehnt:', ''),
          'Abgelehnt ', ''),
        'Refusé :', ''),
      'Rifiutato:', '')
  )
  ```
- Rename → `Cleaned_Subject`.

### 5. `Get_DEX_Event` (SharePoint Get items)

- **SharePoint — Get items**.
- **Site Address:** `DOL-c-DE-EventExperiencePlatform`.
- **List Name:** `DEX_Events`.
- **Show advanced options** → **Filter Query (fx):**
  ```
  concat('Title eq ''', replace(outputs('Cleaned_Subject'), '''', ''''''), '''')
  ```
- **Top Count:** `1`.
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
- **Uri (fx):**
  ```
  concat(
    '_api/web/lists/getbytitle(''Teilnehmer'')/items?$filter=ParticipantEmail eq ''',
    replace(triggerOutputs()?['body/from'], '''', ''''''),
    ''' and Status ne ''Abgemeldet''&$top=1&$select=Id,Status'
  )
  ```
- **Headers:**
  - `Accept: application/json;odata=nometadata`
- Rename → `Get_Teilnehmer_Entry`.

### 8. Condition `Still_Registered`

- Linke Seite **(fx):** `length(body('Get_Teilnehmer_Entry')?['value'])`
- Operator: `is greater than`
- Rechte Seite: `0`.
- Rename → `Still_Registered`.

### 9. `Get_Existing_Reminder` (SharePoint Get items, im Still_Registered/yes)

- **SharePoint — Get items**.
- **Site Address:** `DOL-c-DE-EventExperiencePlatform`.
- **List Name:** `DEX_Emails`.
- **Filter Query (fx):**
  ```
  concat(
    'EmailType eq ''OutlookDeclineReminder'' and Recipient eq ''',
    replace(triggerOutputs()?['body/from'], '''', ''''''),
    ''' and EventId eq ''',
    first(outputs('Get_DEX_Event')?['body/value'])?['ID'],
    ''''
  )
  ```
- **Top Count:** `1`.
- Rename → `Get_Existing_Reminder`.

### 10. Condition `No_Reminder_Yet`

- Linke Seite **(fx):** `length(outputs('Get_Existing_Reminder')?['body/value'])`
- Operator: `is equal to`
- Rechte Seite: `0`.
- Rename → `No_Reminder_Yet`.

### 11. `Create_Reminder_Queue_Item` (SharePoint Create item, im No_Reminder_Yet/yes)

- **SharePoint — Create item**.
- **Site Address:** `DOL-c-DE-EventExperiencePlatform`.
- **List Name:** `DEX_Emails`.
- **Title (fx):** `concat('Outlook-Abmeldung-Reminder: ', first(outputs('Get_DEX_Event')?['body/value'])?['Title'])`
- **Recipient (fx):** `triggerOutputs()?['body/from']`
- **RecipientName (fx):** `triggerOutputs()?['body/from']` (kein Display-Name im Mail-Trigger verfügbar)
- **EmailType Value:** `OutlookDeclineReminder`.
- **EventTitle (fx):** `first(outputs('Get_DEX_Event')?['body/value'])?['Title']`
- **EventId (fx):** `first(outputs('Get_DEX_Event')?['body/value'])?['ID']`
- **Status Value:** `Pending`.
- **Body:** leer lassen (DEX_SEND_MAIL füllt den Body aus dem Template).
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
- Navigiert zu `my-events` mit `selectedEventId` + Intent `auto-cancel`
- `MyEventsPage.tsx` prüft den Intent, scrollt zur Event-Karte und öffnet den Abmelde-Bestätigungsdialog automatisch
- User muss aktiv auf "Abmeldung bestätigen" klicken — der Deep-Link cancelt NICHT direkt, damit Missbrauch ausgeschlossen ist (User muss eingeloggt sein und ist dann nur in der Lage, SEINE EIGENE Registrierung zu canceln)

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

**ROOT Condition `Is_Decline_Mail` + alle Kinder (If):**
```json
{
  "type": "If",
  "expression": {
    "and": [
      {
        "equals": [
          "@or(\n  startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'declined:'),\n  startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'declined '),\n  startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'abgelehnt:'),\n  startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'abgelehnt '),\n  startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'refusé'),\n  startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'rifiutato:')\n)",
          "@true"
        ]
      }
    ]
  },
  "actions": {
    "Cleaned_Subject": {
      "type": "Compose",
      "inputs": "@trim(\n  replace(\n    replace(\n      replace(\n        replace(\n          replace(\n            replace(triggerOutputs()?['body/subject'], 'Declined:', ''),\n            'Declined ', ''),\n          'Abgelehnt:', ''),\n        'Abgelehnt ', ''),\n      'Refusé :', ''),\n    'Rifiutato:', '')\n)"
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
      "runAfter": { "Cleaned_Subject": ["Succeeded"] }
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
              "parameters/uri": "@concat(\n  '_api/web/lists/getbytitle(''Teilnehmer'')/items?$filter=ParticipantEmail eq ''',\n  replace(triggerOutputs()?['body/from'], '''', ''''''),\n  ''' and Status ne ''Abgemeldet''&$top=1&$select=Id,Status'\n)",
              "parameters/headers": { "Accept": "application/json;odata=nometadata" }
            },
            "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "HttpRequest" }
          }
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
                  "$filter": "@concat(\n  'EmailType eq ''OutlookDeclineReminder'' and EventId eq ''',\n  first(outputs('Get_DEX_Event')?['body/value'])?['ID'],\n  ''''\n)",
                  "$top": 20
                },
                "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "GetItems" }
              }
            },
            "Filter_By_Recipient": {
              "type": "Query",
              "inputs": {
                "from": "@body('Get_Existing_Reminder')?['value']",
                "where": "@contains(concat(';', replace(item()?['Recipient'], ' ', ''), ';'), concat(';', triggerOutputs()?['body/from'], ';'))"
              },
              "runAfter": { "Get_Existing_Reminder": ["Succeeded"] }
            },
            "No_Reminder_Yet": {
              "type": "If",
              "expression": { "and": [ { "equals": ["@length(body('Filter_By_Recipient'))", 0] } ] },
              "actions": {
                "Create_Reminder_Queue_Item": {
                  "type": "OpenApiConnection",
                  "inputs": {
                    "parameters": {
                      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                      "table": "57aa0840-df98-41ae-a39b-323c0b80ae3b",
                      "item/Title": "@concat('Outlook-Abmeldung-Reminder: ', first(outputs('Get_DEX_Event')?['body/value'])?['Title'])",
                      "item/Recipient": "@triggerOutputs()?['body/from']",
                      "item/RecipientName": "@triggerOutputs()?['body/from']",
                      "item/EmailType/Value": "OutlookDeclineReminder",
                      "item/EventTitle": "@first(outputs('Get_DEX_Event')?['body/value'])?['Title']",
                      "item/Status/Value": "Pending",
                      "item/EventId": "@first(outputs('Get_DEX_Event')?['body/value'])?['ID']"
                    },
                    "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "PostItem" }
                  }
                }
              },
              "else": { "actions": {} },
              "runAfter": { "Filter_By_Recipient": ["Succeeded"] }
            }
          },
          "else": { "actions": {} },
          "runAfter": { "Get_Teilnehmer_Entry": ["Succeeded"] }
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
