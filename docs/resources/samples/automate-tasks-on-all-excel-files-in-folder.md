---
title: Ausführen eines Scripts für alle Excel-Dateien in einem Ordner
description: Erfahren Sie, wie Sie ein Skript für alle Excel Dateien in einem Ordner auf OneDrive for Business ausführen.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: fb9a4deb01b52ef031cb1ba3400bd6f10de9d9f5
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545791"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="7a725-103">Ausführen eines Scripts für alle Excel-Dateien in einem Ordner</span><span class="sxs-lookup"><span data-stu-id="7a725-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="7a725-104">Dieses Projekt führt eine Reihe von Automatisierungsaufgaben für alle Dateien aus, die sich in einem Ordner auf OneDrive for Business befinden.</span><span class="sxs-lookup"><span data-stu-id="7a725-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="7a725-105">Es könnte auch auf einem SharePoint Ordner verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="7a725-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="7a725-106">Es führt Berechnungen für die Excel Dateien durch, fügt Formatierungen hinzu und fügt einen Kommentar ein, der einen Kollegen [@mentions.](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)</span><span class="sxs-lookup"><span data-stu-id="7a725-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

<span data-ttu-id="7a725-107">Laden Sie die Datei <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>herunter, extrahieren Sie die Dateien in einen Ordner mit dem Titel **Sales,** der in diesem Beispiel verwendet wird, und probieren Sie es selbst aus!</span><span class="sxs-lookup"><span data-stu-id="7a725-107">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="7a725-108">Beispielcode: Hinzufügen von Formatierungen und Einfügen von Kommentaren</span><span class="sxs-lookup"><span data-stu-id="7a725-108">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="7a725-109">Dies ist das Skript, das für jede einzelne Arbeitsmappe ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="7a725-109">This is the script that runs on each individual workbook.</span></span>

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="7a725-110">Power Automate-Flow: Führen Sie das Skript für jede Arbeitsmappe im Ordner aus.</span><span class="sxs-lookup"><span data-stu-id="7a725-110">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="7a725-111">Dieser Flow führt das Skript für jede Arbeitsmappe im Ordner "Verkauf" aus.</span><span class="sxs-lookup"><span data-stu-id="7a725-111">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="7a725-112">Erstellen Sie einen neuen **Instant Cloud-Flow**.</span><span class="sxs-lookup"><span data-stu-id="7a725-112">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="7a725-113">Wählen Sie **Manuell auslösen einen Flow** aus und drücken Sie **Erstellen**.</span><span class="sxs-lookup"><span data-stu-id="7a725-113">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="7a725-114">Fügen Sie einen **neuen Schritt** hinzu, der den **OneDrive for Business** Connector und die Liste der Dateien **in Ordneraktion** verwendet.</span><span class="sxs-lookup"><span data-stu-id="7a725-114">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Der fertige OneDrive for Business-Anschluss in Power Automate":::
1. <span data-ttu-id="7a725-116">Wählen Sie den Ordner "Verkauf" mit den extrahierten Arbeitsmappen aus.</span><span class="sxs-lookup"><span data-stu-id="7a725-116">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="7a725-117">Um sicherzustellen, dass nur Arbeitsmappen ausgewählt sind, wählen Sie **Neuer Schritt** aus, wählen Sie dann **Bedingung** aus und legen Sie die folgenden Werte fest:</span><span class="sxs-lookup"><span data-stu-id="7a725-117">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="7a725-118">**Name** (der OneDrive Dateinamenwert)</span><span class="sxs-lookup"><span data-stu-id="7a725-118">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="7a725-119">"endet mit"</span><span class="sxs-lookup"><span data-stu-id="7a725-119">"ends with"</span></span>
    1. <span data-ttu-id="7a725-120">"xlsx".</span><span class="sxs-lookup"><span data-stu-id="7a725-120">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Der Power Automate Bedingungsblock, der nachfolgende Aktionen auf jede Datei anwendet.":::
1. <span data-ttu-id="7a725-122">Fügen Sie unter **if yes-Zweig** den **Excel Online-Connector (Business)** mit der Aktion **Skript ausführen** hinzu.</span><span class="sxs-lookup"><span data-stu-id="7a725-122">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="7a725-123">Verwenden Sie die folgenden Werte für die Aktion:</span><span class="sxs-lookup"><span data-stu-id="7a725-123">Use the following values for the action:</span></span>
    1. <span data-ttu-id="7a725-124">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="7a725-124">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="7a725-125">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="7a725-125">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="7a725-126">**Datei**: **ID** (der OneDrive Datei-ID-Wert)</span><span class="sxs-lookup"><span data-stu-id="7a725-126">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="7a725-127">**Skript**: Ihr Skriptname</span><span class="sxs-lookup"><span data-stu-id="7a725-127">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Der abgeschlossene Excel Online-Connector (Business) in Power Automate":::
1. <span data-ttu-id="7a725-129">Speichern Sie den Flow und probieren Sie ihn aus.</span><span class="sxs-lookup"><span data-stu-id="7a725-129">Save the flow and try it out.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="7a725-130">Schulungsvideo: Ausführen eines Skripts für alle Excel Dateien in einem Ordner</span><span class="sxs-lookup"><span data-stu-id="7a725-130">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="7a725-131">[Sehen Sie Sudhi Ramamurthy zu Fuß durch dieses Beispiel auf YouTube](https://youtu.be/xMg711o7k6w).</span><span class="sxs-lookup"><span data-stu-id="7a725-131">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
