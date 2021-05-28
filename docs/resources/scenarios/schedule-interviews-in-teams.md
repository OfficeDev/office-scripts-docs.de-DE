---
title: Planen von Interviews in Teams
description: Erfahren Sie, wie Sie Office Skripts verwenden, um eine Teams aus Excel senden.
ms.date: 05/25/2021
localization_priority: Normal
ms.openlocfilehash: f93d9ceca6603ddb9e7123a393787fcf54597cca
ms.sourcegitcommit: 339ecbb9914d54f919e3475018888fb5d00abe89
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/28/2021
ms.locfileid: "52697784"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a>Office Skriptbeispielszenario: Planen von Teams

In diesem Szenario planen Sie als Personalreferent Interviewbesprechungen mit Kandidaten in Teams. Sie verwalten den Interviewzeitplan der Kandidaten in einer Excel Datei. Sie müssen die Besprechungs-Teams an den Kandidaten und die Interviewer senden. Anschließend müssen Sie die Excel mit der Bestätigung aktualisieren, dass Teams gesendet wurden.

Die Lösung besteht aus drei Schritten, die in einem einzigen Power Automate werden.

1. Ein Skript extrahiert Daten aus einer Tabelle und gibt ein Array von Objekten als JSON-Daten zurück.
1. Die Daten werden dann an die Teams **Erstellen** Teams Besprechungsaktion zum Senden von Einladungen gesendet.
1. Dieselben JSON-Daten werden an ein anderes Skript gesendet, um den Status der Einladung zu aktualisieren.

## <a name="scripting-skills-covered"></a>Abgedeckte Skriptkenntnisse

* Power Automate Flüsse
* Teams Integration
* Tabellenparsierung

## <a name="sample-excel-file"></a>Beispieldatei Excel Datei

Laden Sie die Datei <a href="hr-schedule.xlsx">hr-schedule.xlsx, </a> die in dieser Lösung verwendet wird, herunter, und testen Sie sie selbst! Achten Sie darauf, mindestens eine der E-Mail-Adressen so zu ändern, dass Sie eine Einladung erhalten.

## <a name="sample-code-extract-table-data-to-schedule-invites"></a>Beispielcode: Extrahieren von Tabellendaten zum Planen von Einladungen

Nennen Sie dieses Skript **Zeitplanungsinterviews** für den Fluss.

```TypeScript
function main(workbook: ExcelScript.Workbook): InterviewInvite[] {
  const MEETING_DURATION = workbook.getWorksheet("Constants").getRange("B1").getValue() as number;
  const MESSAGE_TEMPLATE = workbook.getWorksheet("Constants").getRange("B2").getValue() as string;

  // Get the interview candidate information.
  const sheet = workbook.getWorksheet("Interviews");
  const table = sheet.getTables()[0];
  const dataRows = table.getRangeBetweenHeaderAndTotal().getValues();

  // Convert the table rows into InterviewInvite objects for the flow.
  let invites: InterviewInvite[] = [];
  dataRows.forEach((row) => {
    const inviteSent = row[1] as boolean;
    if (!inviteSent) {
      const startTime = new Date(Math.round(((row[6] as number) - 25569) * 86400 * 1000));
      const finishTime = new Date(startTime.getTime() + MEETING_DURATION * 60 * 1000);
      const candidateName = row[2] as string;
      const interviewerName = row[4] as string;

      invites.push({
        ID: row[0] as string,
        Candidate: candidateName,
        CandidateEmail: row[3] as string,
        Interviewer: row[4] as string,
        InterviewerEmail: row[5] as string,
        StartTime: startTime.toISOString(),
        FinishTime: finishTime.toISOString(),
        Message: generateInviteMessage(MESSAGE_TEMPLATE, candidateName, interviewerName)
      });
    }    
  });

  console.log(JSON.stringify(invites));
  return invites;
}

function generateInviteMessage(
  messageTemplate: string,
   candidate: string,
   interviewer: string) : string {
  return messageTemplate.replace("_Candidate_", candidate).replace("_Interviewer_", interviewer);
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-code-mark-rows-as-invited"></a>Beispielcode: Markieren von Zeilen als eingeladen

Benennen Sie dieses Skript **Record Sent Invites** für den Fluss.

```TypeScript
function main(workbook: ExcelScript.Workbook, invites: InterviewInvite[]) {
  const table = workbook.getWorksheet("Interviews").getTables()[0];

  // Get the ID and Invite Sent columns from the table.
  const idColumn = table.getColumnByName("ID");
  const idRange = idColumn.getRangeBetweenHeaderAndTotal().getValues();
  const inviteSentColumn = table.getColumnByName("Invite Sent?");

  const dataRowCount = idRange.length;

  // Find matching IDs to mark the correct row.
  for (let row = 0; row < dataRowCount; row++){
    let inviteSent = invites.find((invite) => {
      return invite.ID == idRange[row][0] as string;
    });

    if (inviteSent) {
      inviteSentColumn.getRangeBetweenHeaderAndTotal().getCell(row, 0).setValue(true);
      console.log(`Invite for ${inviteSent.Candidate} has been sent.`);
    }
  } 
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a>Beispielablauf: Führen Sie die Skripts für die Interviewplanung aus, und senden Teams Besprechungen

1. Erstellen Sie einen neuen **Instant Cloud Flow**.
1. Wählen **Sie Manuellen Fluss auslösen aus,** und drücken Sie **die Create -Taste.**
1. Fügen Sie einen **neuen Schritt hinzu,** der **den Excel Online (Business)-Connector** und die **Skriptaktion ausführen** verwendet. Schließen Sie den Connector mit den folgenden Werten ab.
    1. **Location**: OneDrive for Business
    1. **Document Library**: OneDrive
    1. **Datei**: hr-interviews.xlsx *(über den Dateibrowser ausgewählt)*
    1. **Skript**: Planen von Interviews :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="Screenshot des abgeschlossenen Excel Online (Business)-Connectors"::: zum Erhalten von Interviewdaten aus der Arbeitsmappe in Power Automate
1. Fügen Sie einen **Neuen Schritt hinzu,** der die Besprechungsaktion **erstellen Teams** verwendet. Wenn Sie dynamischen Inhalt aus dem Excel auswählen, wird für Ihren Fluss ein **Apply to each** block generiert. Schließen Sie den Connector mit den folgenden Werten ab.
    1. **Kalender-ID**: Kalender
    1. **Betreff**: Contoso-Interview
    1. **Message**: **Message** (the Excel value)
    1. **Zeitzone :** Pacific Standard Time
    1. **Startzeit**: **StartTime** (Excel Wert)
    1. **Endzeit**: **FinishTime** (der Excel Wert)
    1. **Erforderliche Teilnehmer:** **CandidateEmail** ; **InterviewerEmail** (Excel Werte) Screenshot des abgeschlossenen Teams-Connectors zum Planen :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="von Besprechungen in Power Automate":::
1. Fügen Sie im gleichen **"Auf jeden** Block anwenden" einen **weiteren Excel Online (Business)-Connector** mit der **Skriptaktion Ausführen** hinzu. Verwenden Sie die folgenden Werte.
    1. **Location**: OneDrive for Business
    1. **Document Library**: OneDrive
    1. **Datei**: hr-interviews.xlsx *(über den Dateibrowser ausgewählt)*
    1. **Skript**: Gesendete Einladungen aufzeichnen
    1. **invites**: **result** (the Excel value) Screenshot des abgeschlossenen :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="Excel Online (Business)-Connectors,"::: um zu aufzeichnen, dass Einladungen in der Power Automate
1. Speichern Sie den Fluss, und testen Sie ihn.

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a>Schulungsvideo: Senden Teams Besprechung aus Excel Daten

[Sehen Sie sich an, wie Sudhi Ramamurthy eine Version dieses Beispiels auf YouTube durchschaut.](https://youtu.be/HyBdx52NOE8) Seine Version verwendet ein stabileres Skript, das sich ändernde Spalten und veraltete Besprechungszeiten behandelt.
