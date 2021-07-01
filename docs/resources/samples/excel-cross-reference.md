---
title: Querverweisen Excel Dateien mit Power Automate
description: Erfahren Sie, wie Sie Office Skripts und Power Automate zum Querverweisen und Formatieren einer Excel Datei verwenden.
ms.date: 06/25/2021
localization_priority: Normal
ms.openlocfilehash: 89c4a5fa5dcff21681fa20cd4118447d39d9b6da
ms.sourcegitcommit: a063b3faf6c1b7c294bd6a73e46845b352f2a22d
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/29/2021
ms.locfileid: "53202869"
---
# <a name="cross-reference-excel-files-with-power-automate"></a><span data-ttu-id="82af4-103">Querverweisen Excel Dateien mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="82af4-103">Cross-reference Excel files with Power Automate</span></span>

<span data-ttu-id="82af4-104">Diese Lösung zeigt, wie Daten zwischen zwei Excel Dateien verglichen werden, um Abweichungen zu finden.</span><span class="sxs-lookup"><span data-stu-id="82af4-104">This solution shows how to compare data across two Excel files to find discrepancies.</span></span> <span data-ttu-id="82af4-105">Es verwendet Office Skripts, um Daten zu analysieren und Power Automate, um zwischen den Arbeitsmappen zu kommunizieren.</span><span class="sxs-lookup"><span data-stu-id="82af4-105">It uses Office Scripts to analyze data and Power Automate to communicate between the workbooks.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="82af4-106">Beispielszenario</span><span class="sxs-lookup"><span data-stu-id="82af4-106">Example scenario</span></span>

<span data-ttu-id="82af4-107">Sie sind ein Ereignis-Coordinator, der Referenten für anstehende Konferenzen plant.</span><span class="sxs-lookup"><span data-stu-id="82af4-107">You're an event coordinator who is scheduling speakers for upcoming conferences.</span></span> <span data-ttu-id="82af4-108">Sie behalten die Ereignisdaten in einer Kalkulationstabelle und die Lautsprecherregistrierungen in einer anderen.</span><span class="sxs-lookup"><span data-stu-id="82af4-108">You keep the event data in one spreadsheet and the speaker registrations in another.</span></span> <span data-ttu-id="82af4-109">Um sicherzustellen, dass die beiden Arbeitsmappen synchronisiert werden, verwenden Sie einen Fluss mit Office Skripts, um potenzielle Probleme hervorzuheben.</span><span class="sxs-lookup"><span data-stu-id="82af4-109">To ensure the two workbooks are kept in sync, you use a flow with Office Scripts to highlight any potential problems.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="82af4-110">Beispieldateien Excel</span><span class="sxs-lookup"><span data-stu-id="82af4-110">Sample Excel files</span></span>

<span data-ttu-id="82af4-111">Laden Sie die folgenden Dateien herunter, die in dieser Lösung verwendet werden, um sie selbst auszuprobieren!</span><span class="sxs-lookup"><span data-stu-id="82af4-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="82af4-112"><a href="event-data.xlsx">event-data.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="82af4-112"><a href="event-data.xlsx">event-data.xlsx</a></span></span>
1. <span data-ttu-id="82af4-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="82af4-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="82af4-114">Beispielcode: Abrufen von Ereignisdaten</span><span class="sxs-lookup"><span data-stu-id="82af4-114">Sample code: Get event data</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): string {
  // Get the first table in the "Keys" worksheet.
  let table = workbook.getWorksheet('Keys').getTables()[0];

  // Get the rows in the event table.
  let range = table.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  // Save each row as an EventData object. This lets them be passed through Power Automate.
  let records: EventData[] = [];
  for (let row of rows) {
    let [eventId, date, location, capacity] = row;
    records.push({
      eventId: eventId as string,
      date: date as number,
      location: location as string,
      capacity: capacity as number
    })
  }

  // Log the event data to the console and return it for a flow.
  let stringResult = JSON.stringify(records);
  console.log(stringResult);
  return stringResult;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="sample-code-validate-speaker-registrations"></a><span data-ttu-id="82af4-115">Beispielcode: Überprüfen von Lautsprecherregistrierungen</span><span class="sxs-lookup"><span data-stu-id="82af4-115">Sample code: Validate speaker registrations</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
  // Get the first table in the "Transactions" worksheet.
  let table = workbook.getWorksheet('Transactions').getTables()[0];

  // Clear the existing formatting in the table.
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Compare the data in the table to the keys passed into the script.
  let keysObject = JSON.parse(keys) as EventData[];
  let speakerSlotsRemaining = keysObject.map(value => value.capacity);
  let overallMatch = true;

  // Iterate over every row looking for differences from the other worksheet.
  let rows = range.getValues();
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let [eventId, date, location, capacity] = row;
    let match = false;

    // Look at each key provided for a matching Event ID.
    for (let keyIndex = 0; keyIndex < keysObject.length; keyIndex++) {
      let event = keysObject[keyIndex];
      if (event.eventId === eventId) {
        match = true;
        speakerSlotsRemaining[keyIndex]--;
        // If there's a match on the event ID, look for things that don't match and highlight them.
        if (event.date !== date) {
          overallMatch = false;
          range.getCell(i, 1).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (event.location !== location) {
          overallMatch = false;
          range.getCell(i, 2).getFormat()
            .getFill()
            .setColor("FFFF00");
        }

        break;
      }
    }

    // If no matching Event ID is found, highlight the Event ID's cell.
    if (!match) {
      overallMatch = false;
      range.getCell(i, 0).getFormat()
        .getFill()
        .setColor("FFFF00");
    }
  }

  

  // Choose a message to send to the user.
  let returnString = "All the data is in the right order.";
  if (overallMatch === false) {
    returnString = "Mismatch found. Data requires your review.";
  } else if (speakerSlotsRemaining.find(remaining => remaining < 0)){
    returnString = "Event potentially overbooked. Please review."
  }

  console.log("Returning: " + returnString);
  return returnString;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a><span data-ttu-id="82af4-116">Power Automate Ablauf: Inkonsistenzen in den Arbeitsmappen überprüfen</span><span class="sxs-lookup"><span data-stu-id="82af4-116">Power Automate flow: Check for inconsistencies across the workbooks</span></span>

<span data-ttu-id="82af4-117">Dieser Fluss extrahiert die Ereignisinformationen aus der ersten Arbeitsmappe und verwendet diese Daten, um die zweite Arbeitsmappe zu überprüfen.</span><span class="sxs-lookup"><span data-stu-id="82af4-117">This flow extracts the event information from the first workbook and uses that data to validate the second workbook.</span></span>

1. <span data-ttu-id="82af4-118">Melden Sie sich [bei Power Automate an,](https://flow.microsoft.com) und erstellen Sie einen neuen **Instant Cloud Flow.**</span><span class="sxs-lookup"><span data-stu-id="82af4-118">Sign into [Power Automate](https://flow.microsoft.com) and create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="82af4-119">Wählen Sie **manuell einen Fluss auslösen** und erstellen drücken. </span><span class="sxs-lookup"><span data-stu-id="82af4-119">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="82af4-120">Fügen Sie einen **neuen Schritt** hinzu, der den connector Excel **Online (Business)** mit der **Skriptaktion ausführen** verwendet.</span><span class="sxs-lookup"><span data-stu-id="82af4-120">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="82af4-121">Verwenden Sie die folgenden Werte für die Aktion:</span><span class="sxs-lookup"><span data-stu-id="82af4-121">Use the following values for the action:</span></span>
    * <span data-ttu-id="82af4-122">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="82af4-122">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="82af4-123">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="82af4-123">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="82af4-124">**Datei:** event-data.xlsx ([ausgewählt mit der Dateiauswahl](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="82af4-124">**File**: event-data.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="82af4-125">**Skript:** Ereignisdaten abrufen</span><span class="sxs-lookup"><span data-stu-id="82af4-125">**Script**: Get event data</span></span>

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="Der fertige Excel Online (Business)-Connector für das erste Skript in Power Automate.":::

1. <span data-ttu-id="82af4-127">Fügen Sie einen zweiten **neuen Schritt** hinzu, der den connector Excel **Online (Business)** mit der **Skriptaktion ausführen** verwendet.</span><span class="sxs-lookup"><span data-stu-id="82af4-127">Add a second **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="82af4-128">Verwenden Sie die folgenden Werte für die Aktion:</span><span class="sxs-lookup"><span data-stu-id="82af4-128">Use the following values for the action:</span></span>
    * <span data-ttu-id="82af4-129">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="82af4-129">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="82af4-130">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="82af4-130">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="82af4-131">**Datei:** speaker-registration.xlsx ([ausgewählt mit der Dateiauswahl](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="82af4-131">**File**: speaker-registration.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="82af4-132">**Skript:** Überprüfen der Lautsprecherregistrierung</span><span class="sxs-lookup"><span data-stu-id="82af4-132">**Script**: Validate speaker registration</span></span>

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="Der fertige Excel Online(Business)-Connector für das zweite Skript in Power Automate.":::
1. <span data-ttu-id="82af4-134">In diesem Beispiel wird Outlook als E-Mail-Client verwendet.</span><span class="sxs-lookup"><span data-stu-id="82af4-134">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="82af4-135">Sie können jeden E-Mail-Connector verwenden, der Power Automate unterstützt.</span><span class="sxs-lookup"><span data-stu-id="82af4-135">You could use any email connector Power Automate supports.</span></span> <span data-ttu-id="82af4-136">Fügen Sie einen **neuen Schritt** hinzu, der den **Office 365 Outlook** Connector und die Aktion **"Senden und E-Mail(V2)"** verwendet.</span><span class="sxs-lookup"><span data-stu-id="82af4-136">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="82af4-137">Verwenden Sie die folgenden Werte für die Aktion:</span><span class="sxs-lookup"><span data-stu-id="82af4-137">Use the following values for the action:</span></span>
    * <span data-ttu-id="82af4-138">**An:** Ihr Test-E-Mail-Konto (oder persönliche E-Mail)</span><span class="sxs-lookup"><span data-stu-id="82af4-138">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="82af4-139">**Betreff:** Ergebnisse der Ereignisüberprüfung</span><span class="sxs-lookup"><span data-stu-id="82af4-139">**Subject**: Event validation results</span></span>
    * <span data-ttu-id="82af4-140">**Body**: result (_dynamic content from Run script **2**_)</span><span class="sxs-lookup"><span data-stu-id="82af4-140">**Body**: result (_dynamic content from **Run script 2**_)</span></span>

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="Der fertige Office 365 Outlook-Connector in Power Automate.":::
1. <span data-ttu-id="82af4-142">Speichern Sie den Flow, und wählen Sie dann **"Testen"** aus, um ihn auszuprobieren. Sie sollten eine E-Mail mit der Meldung "Konflikt gefunden" erhalten.</span><span class="sxs-lookup"><span data-stu-id="82af4-142">Save the flow, then select **Test** to try it out. You should receive an email saying "Mismatch found.</span></span> <span data-ttu-id="82af4-143">Daten erfordern Ihre Überprüfung."</span><span class="sxs-lookup"><span data-stu-id="82af4-143">Data requires your review."</span></span> <span data-ttu-id="82af4-144">Dies weist darauf hin, dass es Unterschiede zwischen Zeilen in **speaker-registrations.xlsx** und Zeilen in **event-data.xlsx** gibt.</span><span class="sxs-lookup"><span data-stu-id="82af4-144">This indicates there are differences between rows in **speaker-registrations.xlsx** and rows in **event-data.xlsx**.</span></span> <span data-ttu-id="82af4-145">Öffnen **Siespeaker-registrations.xlsx,** um mehrere hervorgehobene Zellen anzuzeigen, in denen potenzielle Probleme mit den Eintragen für die Lautsprecherregistrierung auftreten.</span><span class="sxs-lookup"><span data-stu-id="82af4-145">Open **speaker-registrations.xlsx** to see several highlighted cells where there are potential problems with the speaker registration listings.</span></span>
