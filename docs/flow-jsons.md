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
**Letztes Update:** 2026-04-06

Ablauf: Trigger (neues Event) → Config laden (Logo + Default-Bild) → Compose_Logo → Compose_Image → Platzhalter in OutlookBody ersetzen → Outlook-Termin mit HTML-Body erstellen → CalendarLink in DEX_Events speichern

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
      "item/start": "@formatDateTime(triggerBody()?['StartDate'], 'yyyy-MM-ddTHH:mm:ss')",
      "item/end": "@formatDateTime(triggerBody()?['EndDate'], 'yyyy-MM-ddTHH:mm:ss')",
      "item/timeZone": "(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna",
      "item/requiredAttendees": "@triggerBody()?['OrganizerEmail']",
      "item/body": "<p class=\"editor-paragraph\">@{replace(replace(triggerBody()?['OutlookBody'], '{{LOGO_URL}}', outputs('Compose_Logo')), '{{ORB_URL}}', outputs('Compose_Image'))}</p>"
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
**Zweck:** Teilnehmer zu bestehendem Outlook-Termin einladen oder ausladen (via Graph API)
**Letztes Update:** 2026-04-06

Ablauf: Trigger → Event-Details laden → CalendarLink vorhanden? → Outlook-Event per iCalUId finden → Event-ID speichern → Event gefunden? → Bestehende Attendees laden → Einladen/Ausladen → Status=Sent (oder Failed bei Fehler)

**Concurrency:** 1 (sequentielle Verarbeitung, max 100 wartende Runs)

```json
TRIGGER (Neues Item in DEX_Outlook):
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

CHECK_CALENDARLINK (CalendarLink vorhanden?):
{
  "type": "If",
  "expression": {
    "and": [{ "greater": ["@length(coalesce(first(outputs('Get_Event_Details')?['body/value'])?['CalendarLink'], ''))", 0] }]
  },
  "actions": {
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
    "Set_variable (varRealEventId)": {
      "type": "SetVariable",
      "inputs": { "name": "varRealEventId", "value": "@first(body('Find_Outlook_Event')?['value'])?['id']" },
      "runAfter": { "Find_Outlook_Event": ["Succeeded"] }
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
        "Set_variable_1 (var_Attendees)": {
          "type": "SetVariable",
          "inputs": { "name": "var_Attendees", "value": "@body('Get_Existing_Event')?['attendees']" },
          "runAfter": { "Get_Existing_Event": ["Succeeded"] }
        },
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
          },
          "runAfter": { "Set_variable_1": ["Succeeded"] }
        },
        "Set_Failed_1_1 (ActionType fehlgeschlagen)": {
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
          "runAfter": { "Check_ActionType": ["Failed"] }
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
    },
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
