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
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a><span data-ttu-id="1e9be-103">Office Skriptbeispielszenario: Planen von Teams</span><span class="sxs-lookup"><span data-stu-id="1e9be-103">Office Scripts sample scenario: Schedule interviews in Teams</span></span>

<span data-ttu-id="1e9be-104">In diesem Szenario planen Sie als Personalreferent Interviewbesprechungen mit Kandidaten in Teams.</span><span class="sxs-lookup"><span data-stu-id="1e9be-104">In this scenario, you're an HR recruiter scheduling interview meetings with candidates in Teams.</span></span> <span data-ttu-id="1e9be-105">Sie verwalten den Interviewzeitplan der Kandidaten in einer Excel Datei.</span><span class="sxs-lookup"><span data-stu-id="1e9be-105">You manage the interview schedule of candidates in an Excel file.</span></span> <span data-ttu-id="1e9be-106">Sie müssen die Besprechungs-Teams an den Kandidaten und die Interviewer senden.</span><span class="sxs-lookup"><span data-stu-id="1e9be-106">You'll need to send the Teams meeting invite to both the candidate and interviewers.</span></span> <span data-ttu-id="1e9be-107">Anschließend müssen Sie die Excel mit der Bestätigung aktualisieren, dass Teams gesendet wurden.</span><span class="sxs-lookup"><span data-stu-id="1e9be-107">You then need to update the Excel file with the confirmation that Teams meetings have been sent.</span></span>

<span data-ttu-id="1e9be-108">Die Lösung besteht aus drei Schritten, die in einem einzigen Power Automate werden.</span><span class="sxs-lookup"><span data-stu-id="1e9be-108">The solution has three steps that are combined in a single Power Automate flow.</span></span>

1. <span data-ttu-id="1e9be-109">Ein Skript extrahiert Daten aus einer Tabelle und gibt ein Array von Objekten als JSON-Daten zurück.</span><span class="sxs-lookup"><span data-stu-id="1e9be-109">A script extracts data from a table and returns an array of objects as JSON data.</span></span>
1. <span data-ttu-id="1e9be-110">Die Daten werden dann an die Teams **Erstellen** Teams Besprechungsaktion zum Senden von Einladungen gesendet.</span><span class="sxs-lookup"><span data-stu-id="1e9be-110">The data is then sent to the Teams **Create a Teams meeting** action to send invites.</span></span>
1. <span data-ttu-id="1e9be-111">Dieselben JSON-Daten werden an ein anderes Skript gesendet, um den Status der Einladung zu aktualisieren.</span><span class="sxs-lookup"><span data-stu-id="1e9be-111">The same JSON data is sent to another script to update the status of the invitation.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="1e9be-112">Abgedeckte Skriptkenntnisse</span><span class="sxs-lookup"><span data-stu-id="1e9be-112">Scripting skills covered</span></span>

* <span data-ttu-id="1e9be-113">Power Automate Flüsse</span><span class="sxs-lookup"><span data-stu-id="1e9be-113">Power Automate flows</span></span>
* <span data-ttu-id="1e9be-114">Teams Integration</span><span class="sxs-lookup"><span data-stu-id="1e9be-114">Teams integration</span></span>
* <span data-ttu-id="1e9be-115">Tabellenparsierung</span><span class="sxs-lookup"><span data-stu-id="1e9be-115">Table parsing</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="1e9be-116">Beispieldatei Excel Datei</span><span class="sxs-lookup"><span data-stu-id="1e9be-116">Sample Excel file</span></span>

<span data-ttu-id="1e9be-117">Laden Sie die Datei <a href="hr-schedule.xlsx">hr-schedule.xlsx, </a> die in dieser Lösung verwendet wird, herunter, und testen Sie sie selbst!</span><span class="sxs-lookup"><span data-stu-id="1e9be-117">Download the file <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> used in this solution and try it out yourself!</span></span> <span data-ttu-id="1e9be-118">Achten Sie darauf, mindestens eine der E-Mail-Adressen so zu ändern, dass Sie eine Einladung erhalten.</span><span class="sxs-lookup"><span data-stu-id="1e9be-118">Be sure to change at least one of the email addresses so that you receive an invite.</span></span>

## <a name="sample-code-extract-table-data-to-schedule-invites"></a><span data-ttu-id="1e9be-119">Beispielcode: Extrahieren von Tabellendaten zum Planen von Einladungen</span><span class="sxs-lookup"><span data-stu-id="1e9be-119">Sample code: Extract table data to schedule invites</span></span>

<span data-ttu-id="1e9be-120">Nennen Sie dieses Skript **Zeitplanungsinterviews** für den Fluss.</span><span class="sxs-lookup"><span data-stu-id="1e9be-120">Name this script **Schedule Interviews** for the flow.</span></span>

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

## <a name="sample-code-mark-rows-as-invited"></a><span data-ttu-id="1e9be-121">Beispielcode: Markieren von Zeilen als eingeladen</span><span class="sxs-lookup"><span data-stu-id="1e9be-121">Sample code: Mark rows as invited</span></span>

<span data-ttu-id="1e9be-122">Benennen Sie dieses Skript **Record Sent Invites** für den Fluss.</span><span class="sxs-lookup"><span data-stu-id="1e9be-122">Name this script **Record Sent Invites** for the flow.</span></span>

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

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a><span data-ttu-id="1e9be-123">Beispielablauf: Führen Sie die Skripts für die Interviewplanung aus, und senden Teams Besprechungen</span><span class="sxs-lookup"><span data-stu-id="1e9be-123">Sample flow: Run the interview scheduling scripts and send the Teams meetings</span></span>

1. <span data-ttu-id="1e9be-124">Erstellen Sie einen neuen **Instant Cloud Flow**.</span><span class="sxs-lookup"><span data-stu-id="1e9be-124">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="1e9be-125">Wählen **Sie Manuellen Fluss auslösen aus,** und drücken Sie **die Create -Taste.**</span><span class="sxs-lookup"><span data-stu-id="1e9be-125">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="1e9be-126">Fügen Sie einen **neuen Schritt hinzu,** der **den Excel Online (Business)-Connector** und die **Skriptaktion ausführen** verwendet.</span><span class="sxs-lookup"><span data-stu-id="1e9be-126">Add a **New step** that uses the **Excel Online (Business)** connector and the **Run script** action.</span></span> <span data-ttu-id="1e9be-127">Schließen Sie den Connector mit den folgenden Werten ab.</span><span class="sxs-lookup"><span data-stu-id="1e9be-127">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="1e9be-128">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="1e9be-128">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="1e9be-129">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="1e9be-129">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="1e9be-130">**Datei**: hr-interviews.xlsx *(über den Dateibrowser ausgewählt)*</span><span class="sxs-lookup"><span data-stu-id="1e9be-130">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. **Skript**: Planen von Interviews :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="Screenshot des abgeschlossenen Excel Online (Business)-Connectors"::: zum Erhalten von Interviewdaten aus der Arbeitsmappe in Power Automate
1. <span data-ttu-id="1e9be-132">Fügen Sie einen **Neuen Schritt hinzu,** der die Besprechungsaktion **erstellen Teams** verwendet.</span><span class="sxs-lookup"><span data-stu-id="1e9be-132">Add a **New step** that uses the **Create a Teams meeting** action.</span></span> <span data-ttu-id="1e9be-133">Wenn Sie dynamischen Inhalt aus dem Excel auswählen, wird für Ihren Fluss ein **Apply to each** block generiert.</span><span class="sxs-lookup"><span data-stu-id="1e9be-133">As you select dynamic content from the Excel connector, an **Apply to each** block will be generated for your flow.</span></span> <span data-ttu-id="1e9be-134">Schließen Sie den Connector mit den folgenden Werten ab.</span><span class="sxs-lookup"><span data-stu-id="1e9be-134">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="1e9be-135">**Kalender-ID**: Kalender</span><span class="sxs-lookup"><span data-stu-id="1e9be-135">**Calendar id**: Calendar</span></span>
    1. <span data-ttu-id="1e9be-136">**Betreff**: Contoso-Interview</span><span class="sxs-lookup"><span data-stu-id="1e9be-136">**Subject**: Contoso Interview</span></span>
    1. <span data-ttu-id="1e9be-137">**Message**: **Message** (the Excel value)</span><span class="sxs-lookup"><span data-stu-id="1e9be-137">**Message**: **Message** (the Excel value)</span></span>
    1. <span data-ttu-id="1e9be-138">**Zeitzone :** Pacific Standard Time</span><span class="sxs-lookup"><span data-stu-id="1e9be-138">**Time zone**: Pacific Standard Time</span></span>
    1. <span data-ttu-id="1e9be-139">**Startzeit**: **StartTime** (Excel Wert)</span><span class="sxs-lookup"><span data-stu-id="1e9be-139">**Start time**: **StartTime** (the Excel value)</span></span>
    1. <span data-ttu-id="1e9be-140">**Endzeit**: **FinishTime** (der Excel Wert)</span><span class="sxs-lookup"><span data-stu-id="1e9be-140">**End time**: **FinishTime** (the Excel value)</span></span>
    1. **Erforderliche Teilnehmer:** **CandidateEmail** ; **InterviewerEmail** (Excel Werte) Screenshot des abgeschlossenen Teams-Connectors zum Planen :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="von Besprechungen in Power Automate":::
1. <span data-ttu-id="1e9be-142">Fügen Sie im gleichen **"Auf jeden** Block anwenden" einen **weiteren Excel Online (Business)-Connector** mit der **Skriptaktion Ausführen** hinzu.</span><span class="sxs-lookup"><span data-stu-id="1e9be-142">In the same **Apply to each** block, add another **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="1e9be-143">Verwenden Sie die folgenden Werte.</span><span class="sxs-lookup"><span data-stu-id="1e9be-143">Use the following values.</span></span>
    1. <span data-ttu-id="1e9be-144">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="1e9be-144">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="1e9be-145">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="1e9be-145">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="1e9be-146">**Datei**: hr-interviews.xlsx *(über den Dateibrowser ausgewählt)*</span><span class="sxs-lookup"><span data-stu-id="1e9be-146">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. <span data-ttu-id="1e9be-147">**Skript**: Gesendete Einladungen aufzeichnen</span><span class="sxs-lookup"><span data-stu-id="1e9be-147">**Script**: Record Sent Invites</span></span>
    1. **invites**: **result** (the Excel value) Screenshot des abgeschlossenen :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="Excel Online (Business)-Connectors,"::: um zu aufzeichnen, dass Einladungen in der Power Automate
1. <span data-ttu-id="1e9be-149">Speichern Sie den Fluss, und testen Sie ihn.</span><span class="sxs-lookup"><span data-stu-id="1e9be-149">Save the flow and try it out.</span></span>

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a><span data-ttu-id="1e9be-150">Schulungsvideo: Senden Teams Besprechung aus Excel Daten</span><span class="sxs-lookup"><span data-stu-id="1e9be-150">Training video: Send a Teams meeting from Excel data</span></span>

<span data-ttu-id="1e9be-151">[Sehen Sie sich an, wie Sudhi Ramamurthy eine Version dieses Beispiels auf YouTube durchschaut.](https://youtu.be/HyBdx52NOE8)</span><span class="sxs-lookup"><span data-stu-id="1e9be-151">[Watch Sudhi Ramamurthy walk through a version of this sample on YouTube](https://youtu.be/HyBdx52NOE8).</span></span> <span data-ttu-id="1e9be-152">Seine Version verwendet ein stabileres Skript, das sich ändernde Spalten und veraltete Besprechungszeiten behandelt.</span><span class="sxs-lookup"><span data-stu-id="1e9be-152">His version uses a more robust script that handles changing columns and obsolete meeting times.</span></span>
