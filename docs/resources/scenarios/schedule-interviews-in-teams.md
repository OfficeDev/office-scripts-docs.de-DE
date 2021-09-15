---
title: Vorstellungsgespräche in Teams planen
description: Erfahren Sie, wie Sie Office Skripts verwenden, um eine Teams Besprechung aus Excel Daten zu senden.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 143b2018824f82c329692909d9a7d1ba10c2485f
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/15/2021
ms.locfileid: "59330870"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a>Office Skripts-Beispielszenario: Planen von Interviews in Teams

In diesem Szenario sind Sie ein Personalberater, der Besprechungen mit Kandidaten in Teams plant. Sie verwalten den Vorstellungsterminplan von Kandidaten in einer Excel-Datei. Sie müssen die Teams Besprechungseinladung sowohl an den Kandidaten als auch an die Interviewer senden. Anschließend müssen Sie die Excel-Datei mit der Bestätigung aktualisieren, dass Teams Besprechungen gesendet wurden.

Die Lösung umfasst drei Schritte, die in einem einzelnen Power Automate-Fluss kombiniert werden.

1. Ein Skript extrahiert Daten aus einer Tabelle und gibt ein Array von Objekten als JSON-Daten zurück.
1. Die Daten werden dann an das Teams **Erstellen einer Teams Besprechungsaktion** zum Senden von Einladungen gesendet.
1. Dieselben JSON-Daten werden an ein anderes Skript gesendet, um den Status der Einladung zu aktualisieren.

## <a name="scripting-skills-covered"></a>Abgedeckte Skriptfähigkeiten

* Power Automate Flüsse
* integration von Teams
* Tabellenparsing

## <a name="sample-excel-file"></a>Beispieldatei für Excel

Laden Sie die in dieser Lösung verwendete Datei <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> herunter, und testen Sie sie selbst! Achten Sie darauf, mindestens eine der E-Mail-Adressen zu ändern, damit Sie eine Einladung erhalten.

## <a name="sample-code-extract-table-data-to-schedule-invites"></a>Beispielcode: Extrahieren von Tabellendaten zum Planen von Einladungen

Fügen Sie dieses Skript zu Ihrer Skriptsammlung hinzu. Nennen Sie es **"Planen von Interviews** für den Fluss".

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

Fügen Sie dieses Skript zu Ihrer Skriptsammlung hinzu. Nennen Sie es **"Gesendete Einladungen** aufzeichnen" für den Fluss.

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

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a>Beispielablauf: Ausführen der Skripts für die Planung von Gesprächen und Senden der Teams Besprechungen

1. Erstellen Sie einen neuen **Instant Cloud Flow.**
1. Klicken Sie **auf "Manuell auslösen" und** wählen Sie **"Erstellen"** aus.
1. Fügen Sie einen **neuen Schritt** hinzu, der den Connector Excel **Online (Business)** und die **Skriptaktion ausführen** verwendet. Schließen Sie den Connector mit den folgenden Werten ab.
    1. **Location**: OneDrive for Business
    1. **Document Library**: OneDrive
    1. **Datei:** hr-interviews.xlsx *(über den Dateibrowser ausgewählt)*
    1. **Skript:** :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="Screenshot des abgeschlossenen Excel Online (Business)-Connectors zum Abrufen von Interviewdaten aus der Arbeitsmappe in Power Automate.":::
1. Fügen Sie einen **neuen Schritt** hinzu, in dem die Besprechungsaktion **"Teams erstellen"** verwendet wird. Wenn Sie dynamische Inhalte aus dem Excel Connector auswählen, wird für den Fluss ein **Apply-Element** für jeden Block generiert. Schließen Sie den Connector mit den folgenden Werten ab.
    1. **Kalender-ID**: Kalender
    1. **Betreff:** Contoso-Interview
    1. **Message**: **Message** (the Excel value)
    1. **Zeitzone:** Pazifische Standardzeit
    1. **Startzeit:** **StartTime** (der wert Excel)
    1. **Endzeit:** **FinishTime** (der wert Excel)
    1. **Erforderliche Teilnehmer:** **CandidateEmail** ; **InterviewerEmail** (die Excel Werte) :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="Screenshot des abgeschlossenen Teams Connectors zum Planen von Besprechungen in Power Automate.":::
1. Fügen Sie im selben **"Auf jeden** Block anwenden" einen weiteren **Excel Online(Business)-Connector** mit der **Skriptaktion ausführen** hinzu. Verwenden Sie die folgenden Werte.
    1. **Location**: OneDrive for Business
    1. **Document Library**: OneDrive
    1. **Datei:** hr-interviews.xlsx *(über den Dateibrowser ausgewählt)*
    1. **Skript:** Gesendete Einladungen aufzeichnen
    1. **invites**: **result** (the Excel value) :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="screenshot of the completed Excel Online (Business) connector to record that invites have been sent in Power Automate.":::
1. Speichern Sie den Fluss, und testen Sie ihn. Verwenden Sie die Schaltfläche **"Test"** auf der Flow-Editor-Seite, oder führen Sie den Fluss über Ihre Registerkarte **"Meine Flüsse"** aus. Achten Sie darauf, den Zugriff zuzulassen, wenn Sie dazu aufgefordert werden.

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a>Schulungsvideo: Senden einer Teams Besprechung aus Excel Daten

[Sehen Sie sich auf YouTube einen Ausführlichen Überblick über eine Version dieses Beispiels an.](https://youtu.be/HyBdx52NOE8) Seine Version verwendet ein robusteres Skript, das sich ändernde Spalten und veraltete Besprechungszeiten verarbeitet.
