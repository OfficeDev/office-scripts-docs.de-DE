---
title: E-Mail der Bilder eines Excel Und Tabelle
description: Erfahren Sie, wie Sie Office Skripts und Power Automate verwenden, um die Bilder eines Diagramms und einer Tabelle Excel zu extrahieren und per E-Mail zu senden.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 54b6b67a0f211f2dc6c881bab17ff23220619e6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545776"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="f94d7-103">Verwenden Office Skripts und Power Automate, um Bilder eines Diagramms und einer Tabelle per E-Mail zu senden</span><span class="sxs-lookup"><span data-stu-id="f94d7-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="f94d7-104">In diesem Beispiel werden Office Skripts und Power Automate zum Erstellen eines Diagramms verwendet.</span><span class="sxs-lookup"><span data-stu-id="f94d7-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="f94d7-105">Anschließend werden Bilder des Diagramms und seiner Basistabelle per E-Mail gesendet.</span><span class="sxs-lookup"><span data-stu-id="f94d7-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="f94d7-106">Beispielszenario</span><span class="sxs-lookup"><span data-stu-id="f94d7-106">Example scenario</span></span>

* <span data-ttu-id="f94d7-107">Berechnen Sie, um die neuesten Ergebnisse zu erhalten.</span><span class="sxs-lookup"><span data-stu-id="f94d7-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="f94d7-108">Diagramm erstellen.</span><span class="sxs-lookup"><span data-stu-id="f94d7-108">Create chart.</span></span>
* <span data-ttu-id="f94d7-109">Get chart and table images.</span><span class="sxs-lookup"><span data-stu-id="f94d7-109">Get chart and table images.</span></span>
* <span data-ttu-id="f94d7-110">Senden Sie die Bilder per E-Mail Power Automate.</span><span class="sxs-lookup"><span data-stu-id="f94d7-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="f94d7-111">_Eingabedaten_</span><span class="sxs-lookup"><span data-stu-id="f94d7-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Ein Arbeitsblatt mit einer Tabelle mit Eingabedaten":::

<span data-ttu-id="f94d7-113">_Ausgabediagramm_</span><span class="sxs-lookup"><span data-stu-id="f94d7-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="Das vom Kunden erstellte Spaltendiagramm mit dem fälligen Betrag":::

<span data-ttu-id="f94d7-115">_E-Mails, die über den Power Automate empfangen wurden_</span><span class="sxs-lookup"><span data-stu-id="f94d7-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="Die vom Fluss gesendete E-Mail mit dem Excel, das in den Textkörper eingebettet ist":::

## <a name="solution"></a><span data-ttu-id="f94d7-117">Lösung</span><span class="sxs-lookup"><span data-stu-id="f94d7-117">Solution</span></span>

<span data-ttu-id="f94d7-118">Diese Lösung besteht aus zwei Teilen:</span><span class="sxs-lookup"><span data-stu-id="f94d7-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="f94d7-119">Ein Office Skript zum Berechnen und Extrahieren Excel Diagramms und der Tabelle</span><span class="sxs-lookup"><span data-stu-id="f94d7-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="f94d7-120">Ein Power Automate zum Aufrufen des Skripts und zum Senden der Ergebnisse per E-Mail.</span><span class="sxs-lookup"><span data-stu-id="f94d7-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="f94d7-121">Ein Beispiel dazu finden Sie unter Erstellen eines automatisierten [Workflows mit Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="f94d7-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="f94d7-122">Beispielcode: Berechnen und Extrahieren Excel Diagramm und Tabelle</span><span class="sxs-lookup"><span data-stu-id="f94d7-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="f94d7-123">Das folgende Skript berechnet und extrahiert ein Excel Und Tabelle.</span><span class="sxs-lookup"><span data-stu-id="f94d7-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="f94d7-124">Laden Sie die Beispieldatei <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> und verwenden Sie sie mit diesem Skript, um sie selbst auszuprobieren!</span><span class="sxs-lookup"><span data-stu-id="f94d7-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  // Recalculate the workbook to ensure all tables and charts are updated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  // Get the data from the "InvoiceAmounts" table.
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  // Get only the "Customer Name" and "Amount due" columns, then remove the "Total" row.
  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  // Delete the "ChartSheet" worksheet if it's present, then recreate it.
  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');

  // Add the selected data to the new worksheet.
  const targetRange = chartSheet.getRange('A1').getResizedRange(selectColumns.length-1, selectColumns[0].length-1);
  targetRange.setValues(selectColumns);

  // Insert the chart on sheet 'ChartSheet' at cell "D1".
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');

  // Get images of the chart and table, then return them for a Power Automate flow.
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {chartImage, tableImage};
}

// The interface for table and chart images.
interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="f94d7-125">Power Automate: E-Mail-E-Mail des Diagramms und der Tabellenbilder</span><span class="sxs-lookup"><span data-stu-id="f94d7-125">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="f94d7-126">In diesem Fluss wird das Skript ausgeführt und die zurückgegebenen Bilder per E-Mail gesendet.</span><span class="sxs-lookup"><span data-stu-id="f94d7-126">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="f94d7-127">Erstellen Sie einen neuen **Instant Cloud Flow**.</span><span class="sxs-lookup"><span data-stu-id="f94d7-127">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="f94d7-128">Wählen **Sie Manuellen Fluss auslösen aus,** und drücken Sie **die Create -Taste.**</span><span class="sxs-lookup"><span data-stu-id="f94d7-128">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="f94d7-129">Fügen Sie einen **neuen Schritt** hinzu, der den Excel **Online (Business)-Connector** mit der **Skriptaktion Ausführen** verwendet.</span><span class="sxs-lookup"><span data-stu-id="f94d7-129">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="f94d7-130">Verwenden Sie die folgenden Werte für die Aktion:</span><span class="sxs-lookup"><span data-stu-id="f94d7-130">Use the following values for the action:</span></span>
    * <span data-ttu-id="f94d7-131">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="f94d7-131">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="f94d7-132">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="f94d7-132">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="f94d7-133">**Datei**: Ihre Arbeitsmappe ([ausgewählt mit der Datei-Auswahl](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="f94d7-133">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="f94d7-134">**Skript**: Ihr Skriptname</span><span class="sxs-lookup"><span data-stu-id="f94d7-134">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Der abgeschlossene Excel Online (Business)-Connector in Power Automate":::
1. <span data-ttu-id="f94d7-136">In diesem Beispiel wird Outlook als E-Mail-Client verwendet.</span><span class="sxs-lookup"><span data-stu-id="f94d7-136">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="f94d7-137">Sie können beliebige E-Mail-Power Automate verwenden, aber bei den restlichen Schritten wird davon ausgegangen, dass Sie Outlook.</span><span class="sxs-lookup"><span data-stu-id="f94d7-137">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="f94d7-138">Fügen Sie einen **Neuen Schritt hinzu,** der den **Office 365 Outlook** und die Aktion Senden **und E-Mail (V2)** verwendet.</span><span class="sxs-lookup"><span data-stu-id="f94d7-138">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="f94d7-139">Verwenden Sie die folgenden Werte für die Aktion:</span><span class="sxs-lookup"><span data-stu-id="f94d7-139">Use the following values for the action:</span></span>
    * <span data-ttu-id="f94d7-140">**An**: Ihr Test-E-Mail-Konto (oder persönliche E-Mail)</span><span class="sxs-lookup"><span data-stu-id="f94d7-140">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="f94d7-141">**Betreff**: Please Review Report Data</span><span class="sxs-lookup"><span data-stu-id="f94d7-141">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="f94d7-142">Wählen Sie **für das Feld Text** die Option "Codeansicht" ( ) `</>` aus, und geben Sie Folgendes ein:</span><span class="sxs-lookup"><span data-stu-id="f94d7-142">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

    ```HTML
    <p>Please review the following report data:<br>
    <br>
    Chart:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/chartImage']}"/>
    <br>
    Data:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/tableImage']}"/>
    <br>
    </p>
    ```

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Der fertige Office 365 Outlook in Power Automate":::
1. <span data-ttu-id="f94d7-144">Speichern Sie den Fluss, und testen Sie ihn.</span><span class="sxs-lookup"><span data-stu-id="f94d7-144">Save the flow and try it out.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="f94d7-145">Schulungsvideo: Extrahieren und E-Mail-Bilder von Diagramm und Tabelle</span><span class="sxs-lookup"><span data-stu-id="f94d7-145">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="f94d7-146">[Sehen Sie sich an, wie Sudhi Ramamurthy dieses Beispiel auf YouTube durchspazieren.](https://youtu.be/152GJyqc-Kw)</span><span class="sxs-lookup"><span data-stu-id="f94d7-146">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
