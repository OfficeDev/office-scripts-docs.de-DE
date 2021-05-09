---
title: Ausführen eines Scripts für alle Excel-Dateien in einem Ordner
description: Erfahren Sie, wie Sie ein Skript für alle Excel in einem Ordner auf einem OneDrive for Business.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: cfe603f3b7fa0ffc27aa3478b2f54788ad645b3f
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285808"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="4b1bb-103">Ausführen eines Scripts für alle Excel-Dateien in einem Ordner</span><span class="sxs-lookup"><span data-stu-id="4b1bb-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="4b1bb-104">Dieses Projekt führt eine Reihe von Automatisierungsaufgaben für alle Dateien aus, die sich in einem Ordner auf einem OneDrive for Business.</span><span class="sxs-lookup"><span data-stu-id="4b1bb-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="4b1bb-105">Es kann auch für einen Ordner SharePoint werden.</span><span class="sxs-lookup"><span data-stu-id="4b1bb-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="4b1bb-106">Es führt Berechnungen für die Excel aus, fügt Formatierungen hinzu und fügt einen Kommentar ein, der [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) kollegen.</span><span class="sxs-lookup"><span data-stu-id="4b1bb-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

<span data-ttu-id="4b1bb-107">Laden Sie die Datei <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extrahieren Sie die Dateien in einen Ordner mit dem Titel **Sales,** der in diesem Beispiel verwendet wird, und testen Sie sie selbst!</span><span class="sxs-lookup"><span data-stu-id="4b1bb-107">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="4b1bb-108">Beispielcode: Hinzufügen von Formatierung und Einfügen eines Kommentars</span><span class="sxs-lookup"><span data-stu-id="4b1bb-108">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="4b1bb-109">Dies ist das Skript, das für jede einzelne Arbeitsmappe ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="4b1bb-109">This is the script that runs on each individual workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "Table1" in the workbook.
  let table1 = workbook.getTable("Table1");

  // If the table is empty, end the script.
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }

  // Force the workbook to be completely recalculated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  // Get the "Amount Due" column from the table.
  const amountDueColumn = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueColumn.getRangeBetweenHeaderAndTotal().getValues();

  // Find the highest amount that's due.
  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }

  let highestAmountDue = table1.getColumn("Amount due").getRangeBetweenHeaderAndTotal().getRow(row);

  // Set the fill color to yellow for the cell with the highest value in the "Amount Due" column.
  highestAmountDue
    .getFormat()
    .getFill()
    .setColor("FFFF00");

  // Insert an @mention comment in the cell.
  workbook.addComment(highestAmountDue, {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="4b1bb-110">Power Automate: Ausführen des Skripts für jede Arbeitsmappe im Ordner</span><span class="sxs-lookup"><span data-stu-id="4b1bb-110">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="4b1bb-111">In diesem Fluss wird das Skript für jede Arbeitsmappe im Ordner "Sales" ausgeführt.</span><span class="sxs-lookup"><span data-stu-id="4b1bb-111">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="4b1bb-112">Erstellen Sie einen neuen **Instant Cloud Flow**.</span><span class="sxs-lookup"><span data-stu-id="4b1bb-112">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="4b1bb-113">Wählen **Sie Manuellen Fluss auslösen aus,** und drücken Sie **die Create -Taste.**</span><span class="sxs-lookup"><span data-stu-id="4b1bb-113">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="4b1bb-114">Fügen Sie einen **neuen Schritt hinzu,** der den **OneDrive for Business** und die Aktion Dateien in Ordner **auflisten** verwendet.</span><span class="sxs-lookup"><span data-stu-id="4b1bb-114">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Der fertige OneDrive for Business in Power Automate":::
1. <span data-ttu-id="4b1bb-116">Wählen Sie den Ordner "Vertrieb" mit den extrahierten Arbeitsmappen aus.</span><span class="sxs-lookup"><span data-stu-id="4b1bb-116">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="4b1bb-117">Um sicherzustellen, dass nur Arbeitsmappen ausgewählt sind, wählen Sie **Neuer Schritt** aus, wählen Sie **dann Bedingung** aus, und legen Sie die folgenden Werte sicher:</span><span class="sxs-lookup"><span data-stu-id="4b1bb-117">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="4b1bb-118">**Name** (der OneDrive Dateinamewert)</span><span class="sxs-lookup"><span data-stu-id="4b1bb-118">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="4b1bb-119">"endet mit"</span><span class="sxs-lookup"><span data-stu-id="4b1bb-119">"ends with"</span></span>
    1. <span data-ttu-id="4b1bb-120">"xlsx".</span><span class="sxs-lookup"><span data-stu-id="4b1bb-120">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Der Power Automate-Bedingungsblock, der nachfolgende Aktionen auf jede Datei angewendet":::
1. <span data-ttu-id="4b1bb-122">Fügen Sie **unter If yes** branch den Excel Online **(Business)** mit der **Aktion Skript ausführen (Vorschau)** hinzu.</span><span class="sxs-lookup"><span data-stu-id="4b1bb-122">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script (preview)** action.</span></span> <span data-ttu-id="4b1bb-123">Verwenden Sie die folgenden Werte für die Aktion:</span><span class="sxs-lookup"><span data-stu-id="4b1bb-123">Use the following values for the action:</span></span>
    1. <span data-ttu-id="4b1bb-124">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="4b1bb-124">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="4b1bb-125">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="4b1bb-125">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="4b1bb-126">**Datei**: **ID** (der OneDrive Datei-ID-Wert)</span><span class="sxs-lookup"><span data-stu-id="4b1bb-126">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="4b1bb-127">**Skript**: Ihr Skriptname</span><span class="sxs-lookup"><span data-stu-id="4b1bb-127">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Der abgeschlossene Excel Online (Business)-Connector in Power Automate":::
1. <span data-ttu-id="4b1bb-129">Speichern Sie den Fluss, und testen Sie ihn.</span><span class="sxs-lookup"><span data-stu-id="4b1bb-129">Save the flow and try it out.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="4b1bb-130">Schulungsvideo: Ausführen eines Skripts für alle Excel in einem Ordner</span><span class="sxs-lookup"><span data-stu-id="4b1bb-130">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="4b1bb-131">[Sehen Sie sich an, wie Sudhi Ramamurthy dieses Beispiel auf YouTube durchspazieren.](https://youtu.be/xMg711o7k6w)</span><span class="sxs-lookup"><span data-stu-id="4b1bb-131">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
