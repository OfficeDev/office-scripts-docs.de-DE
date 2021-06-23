---
title: Vorstellungsgespräche in Teams planen
description: Erfahren Sie, wie Sie Office Skripts verwenden, um eine Teams Besprechung aus Excel Daten zu senden.
ms.date: 05/25/2021
localization_priority: Normal
ms.openlocfilehash: 66dae536c4a51ff3e028f06bf3aef3c7509d83bb
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074431"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a><span data-ttu-id="2948a-103">Office Skripts-Beispielszenario: Planen von Interviews in Teams</span><span class="sxs-lookup"><span data-stu-id="2948a-103">Office Scripts sample scenario: Schedule interviews in Teams</span></span>

<span data-ttu-id="2948a-104">In diesem Szenario sind Sie ein Personalberater, der Besprechungen mit Kandidaten in Teams plant.</span><span class="sxs-lookup"><span data-stu-id="2948a-104">In this scenario, you're an HR recruiter scheduling interview meetings with candidates in Teams.</span></span> <span data-ttu-id="2948a-105">Sie verwalten den Terminplan der Kandidaten in einer Excel-Datei.</span><span class="sxs-lookup"><span data-stu-id="2948a-105">You manage the interview schedule of candidates in an Excel file.</span></span> <span data-ttu-id="2948a-106">Sie müssen die Teams Besprechungseinladung sowohl an den Kandidaten als auch an die Interviewer senden.</span><span class="sxs-lookup"><span data-stu-id="2948a-106">You'll need to send the Teams meeting invite to both the candidate and interviewers.</span></span> <span data-ttu-id="2948a-107">Anschließend müssen Sie die Excel-Datei mit der Bestätigung aktualisieren, dass Teams Besprechungen gesendet wurden.</span><span class="sxs-lookup"><span data-stu-id="2948a-107">You then need to update the Excel file with the confirmation that Teams meetings have been sent.</span></span>

<span data-ttu-id="2948a-108">Die Lösung umfasst drei Schritte, die in einem einzigen Power Automate-Fluss kombiniert werden.</span><span class="sxs-lookup"><span data-stu-id="2948a-108">The solution has three steps that are combined in a single Power Automate flow.</span></span>

1. <span data-ttu-id="2948a-109">Ein Skript extrahiert Daten aus einer Tabelle und gibt ein Array von Objekten als JSON-Daten zurück.</span><span class="sxs-lookup"><span data-stu-id="2948a-109">A script extracts data from a table and returns an array of objects as JSON data.</span></span>
1. <span data-ttu-id="2948a-110">Die Daten werden dann an das Teams **Erstellen einer Teams Besprechungsaktion** zum Senden von Einladungen gesendet.</span><span class="sxs-lookup"><span data-stu-id="2948a-110">The data is then sent to the Teams **Create a Teams meeting** action to send invites.</span></span>
1. <span data-ttu-id="2948a-111">Dieselben JSON-Daten werden an ein anderes Skript gesendet, um den Status der Einladung zu aktualisieren.</span><span class="sxs-lookup"><span data-stu-id="2948a-111">The same JSON data is sent to another script to update the status of the invitation.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="2948a-112">Abgedeckte Skriptfähigkeiten</span><span class="sxs-lookup"><span data-stu-id="2948a-112">Scripting skills covered</span></span>

* <span data-ttu-id="2948a-113">Power Automate Flüsse</span><span class="sxs-lookup"><span data-stu-id="2948a-113">Power Automate flows</span></span>
* <span data-ttu-id="2948a-114">integration von Teams</span><span class="sxs-lookup"><span data-stu-id="2948a-114">Teams integration</span></span>
* <span data-ttu-id="2948a-115">Tabellenparsing</span><span class="sxs-lookup"><span data-stu-id="2948a-115">Table parsing</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="2948a-116">Beispieldatei für Excel</span><span class="sxs-lookup"><span data-stu-id="2948a-116">Sample Excel file</span></span>

<span data-ttu-id="2948a-117">Laden Sie die in dieser Lösung verwendete Datei <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> herunter, und testen Sie sie selbst!</span><span class="sxs-lookup"><span data-stu-id="2948a-117">Download the file <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> used in this solution and try it out yourself!</span></span> <span data-ttu-id="2948a-118">Achten Sie darauf, mindestens eine der E-Mail-Adressen zu ändern, damit Sie eine Einladung erhalten.</span><span class="sxs-lookup"><span data-stu-id="2948a-118">Be sure to change at least one of the email addresses so that you receive an invite.</span></span>

## <a name="sample-code-extract-table-data-to-schedule-invites"></a><span data-ttu-id="2948a-119">Beispielcode: Extrahieren von Tabellendaten zum Planen von Einladungen</span><span class="sxs-lookup"><span data-stu-id="2948a-119">Sample code: Extract table data to schedule invites</span></span>

<span data-ttu-id="2948a-120">Nennen Sie dieses Skript **"Interviews planen"** für den Fluss.</span><span class="sxs-lookup"><span data-stu-id="2948a-120">Name this script **Schedule Interviews** for the flow.</span></span>

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

## <a name="sample-code-mark-rows-as-invited"></a><span data-ttu-id="2948a-121">Beispielcode: Markieren von Zeilen als eingeladen</span><span class="sxs-lookup"><span data-stu-id="2948a-121">Sample code: Mark rows as invited</span></span>

<span data-ttu-id="2948a-122">Nennen Sie dieses Skript **"Gesendete Einladungen** aufzeichnen" für den Fluss.</span><span class="sxs-lookup"><span data-stu-id="2948a-122">Name this script **Record Sent Invites** for the flow.</span></span>

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

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a><span data-ttu-id="2948a-123">Beispielablauf: Ausführen der Skripts für die Planung von Gesprächen und Senden der Teams Besprechungen</span><span class="sxs-lookup"><span data-stu-id="2948a-123">Sample flow: Run the interview scheduling scripts and send the Teams meetings</span></span>

1. <span data-ttu-id="2948a-124">Erstellen Sie einen neuen **Instant Cloud Flow.**</span><span class="sxs-lookup"><span data-stu-id="2948a-124">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="2948a-125">Wählen Sie **manuell einen Fluss auslösen** und erstellen drücken. </span><span class="sxs-lookup"><span data-stu-id="2948a-125">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="2948a-126">Fügen Sie einen **neuen Schritt** hinzu, der den Connector Excel **Online (Business)** und die **Skriptaktion ausführen** verwendet.</span><span class="sxs-lookup"><span data-stu-id="2948a-126">Add a **New step** that uses the **Excel Online (Business)** connector and the **Run script** action.</span></span> <span data-ttu-id="2948a-127">Schließen Sie den Connector mit den folgenden Werten ab.</span><span class="sxs-lookup"><span data-stu-id="2948a-127">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="2948a-128">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="2948a-128">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="2948a-129">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="2948a-129">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="2948a-130">**Datei:** hr-interviews.xlsx *(über den Dateibrowser ausgewählt)*</span><span class="sxs-lookup"><span data-stu-id="2948a-130">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. **Skript:** :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="Screenshot des abgeschlossenen Excel Online (Business)-Connectors zum Abrufen von Interviewdaten aus der Arbeitsmappe in Power Automate.":::
1. <span data-ttu-id="2948a-132">Fügen Sie einen **neuen Schritt** hinzu, in dem die Besprechungsaktion **"Teams erstellen"** verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="2948a-132">Add a **New step** that uses the **Create a Teams meeting** action.</span></span> <span data-ttu-id="2948a-133">Wenn Sie dynamische Inhalte aus dem Excel Connector auswählen, wird für den Fluss ein **Apply-Element** für jeden Block generiert.</span><span class="sxs-lookup"><span data-stu-id="2948a-133">As you select dynamic content from the Excel connector, an **Apply to each** block will be generated for your flow.</span></span> <span data-ttu-id="2948a-134">Schließen Sie den Connector mit den folgenden Werten ab.</span><span class="sxs-lookup"><span data-stu-id="2948a-134">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="2948a-135">**Kalender-ID**: Kalender</span><span class="sxs-lookup"><span data-stu-id="2948a-135">**Calendar id**: Calendar</span></span>
    1. <span data-ttu-id="2948a-136">**Betreff:** Contoso-Interview</span><span class="sxs-lookup"><span data-stu-id="2948a-136">**Subject**: Contoso Interview</span></span>
    1. <span data-ttu-id="2948a-137">**Message**: **Message** (the Excel value)</span><span class="sxs-lookup"><span data-stu-id="2948a-137">**Message**: **Message** (the Excel value)</span></span>
    1. <span data-ttu-id="2948a-138">**Zeitzone:** Pazifische Standardzeit</span><span class="sxs-lookup"><span data-stu-id="2948a-138">**Time zone**: Pacific Standard Time</span></span>
    1. <span data-ttu-id="2948a-139">**Startzeit:** **StartTime** (der Excel-Wert)</span><span class="sxs-lookup"><span data-stu-id="2948a-139">**Start time**: **StartTime** (the Excel value)</span></span>
    1. <span data-ttu-id="2948a-140">**Endzeit:** **FinishTime** (der wert Excel)</span><span class="sxs-lookup"><span data-stu-id="2948a-140">**End time**: **FinishTime** (the Excel value)</span></span>
    1. **Erforderliche Teilnehmer:** **CandidateEmail** ; **InterviewerEmail** (die Excel Werte) :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="Screenshot des abgeschlossenen Teams Connectors zum Planen von Besprechungen in Power Automate.":::
1. <span data-ttu-id="2948a-142">Fügen Sie im selben **"Auf jeden** Block anwenden" einen weiteren **Excel Online(Business)-Connector** mit der **Skriptaktion ausführen** hinzu.</span><span class="sxs-lookup"><span data-stu-id="2948a-142">In the same **Apply to each** block, add another **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="2948a-143">Verwenden Sie die folgenden Werte.</span><span class="sxs-lookup"><span data-stu-id="2948a-143">Use the following values.</span></span>
    1. <span data-ttu-id="2948a-144">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="2948a-144">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="2948a-145">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="2948a-145">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="2948a-146">**Datei:** hr-interviews.xlsx *(über den Dateibrowser ausgewählt)*</span><span class="sxs-lookup"><span data-stu-id="2948a-146">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. <span data-ttu-id="2948a-147">**Skript:** Gesendete Einladungen aufzeichnen</span><span class="sxs-lookup"><span data-stu-id="2948a-147">**Script**: Record Sent Invites</span></span>
    1. **invites**: **result** (the Excel value) :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="Screenshot des abgeschlossenen Excel Online (Business)-Connectors, um zu erfassen, dass Einladungen in Power Automate gesendet wurden.":::
1. <span data-ttu-id="2948a-149">Speichern Sie den Flow, und testen Sie ihn.</span><span class="sxs-lookup"><span data-stu-id="2948a-149">Save the flow and try it out.</span></span>

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a><span data-ttu-id="2948a-150">Schulungsvideo: Senden einer Teams Besprechung aus Excel Daten</span><span class="sxs-lookup"><span data-stu-id="2948a-150">Training video: Send a Teams meeting from Excel data</span></span>

<span data-ttu-id="2948a-151">[Sehen Sie sich auf YouTube einen Ausführlichen Überblick über eine Version dieses Beispiels an.](https://youtu.be/HyBdx52NOE8)</span><span class="sxs-lookup"><span data-stu-id="2948a-151">[Watch Sudhi Ramamurthy walk through a version of this sample on YouTube](https://youtu.be/HyBdx52NOE8).</span></span> <span data-ttu-id="2948a-152">Seine Version verwendet ein robusteres Skript, das sich ändernde Spalten und veraltete Besprechungszeiten verarbeitet.</span><span class="sxs-lookup"><span data-stu-id="2948a-152">His version uses a more robust script that handles changing columns and obsolete meeting times.</span></span>
