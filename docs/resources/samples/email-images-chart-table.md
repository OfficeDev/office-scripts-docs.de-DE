---
title: Senden einer E-Mail an die Bilder eines Excel Diagramms und einer Tabelle
description: Erfahren Sie, wie Sie Office Skripts und Power Automate verwenden, um die Bilder eines Excel Diagramms und einer Tabelle zu extrahieren und per E-Mail zu senden.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 5eb20025462614d62774ae6c088bdf397dcfb39d
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074592"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="33e23-103">Verwenden von Office Skripts und Power Automate zum Senden von E-Mail-Bildern eines Diagramms und einer Tabelle</span><span class="sxs-lookup"><span data-stu-id="33e23-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="33e23-104">In diesem Beispiel werden Office Skripts und Power Automate zum Erstellen eines Diagramms verwendet.</span><span class="sxs-lookup"><span data-stu-id="33e23-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="33e23-105">Anschließend werden Bilder des Diagramms und seiner Basistabelle e-Mails gesendet.</span><span class="sxs-lookup"><span data-stu-id="33e23-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="33e23-106">Beispielszenario</span><span class="sxs-lookup"><span data-stu-id="33e23-106">Example scenario</span></span>

* <span data-ttu-id="33e23-107">Berechnen Sie, um die neuesten Ergebnisse zu erhalten.</span><span class="sxs-lookup"><span data-stu-id="33e23-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="33e23-108">Diagramm erstellen.</span><span class="sxs-lookup"><span data-stu-id="33e23-108">Create chart.</span></span>
* <span data-ttu-id="33e23-109">Abrufen von Diagramm- und Tabellenbildern.</span><span class="sxs-lookup"><span data-stu-id="33e23-109">Get chart and table images.</span></span>
* <span data-ttu-id="33e23-110">Senden Sie eine E-Mail an die Bilder mit Power Automate.</span><span class="sxs-lookup"><span data-stu-id="33e23-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="33e23-111">_Eingabedaten_</span><span class="sxs-lookup"><span data-stu-id="33e23-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Ein Arbeitsblatt mit einer Tabelle mit Eingabedaten.":::

<span data-ttu-id="33e23-113">_Ausgabediagramm_</span><span class="sxs-lookup"><span data-stu-id="33e23-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="Das vom Kunden erstellte Säulendiagramm mit dem vom Kunden fälligen Betrag.":::

<span data-ttu-id="33e23-115">_E-Mail, die über Power Automate Fluss empfangen wurde_</span><span class="sxs-lookup"><span data-stu-id="33e23-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="Die vom Fluss gesendete E-Mail mit dem Excel diagramm, das in den Textkörper eingebettet ist.":::

## <a name="solution"></a><span data-ttu-id="33e23-117">Lösung</span><span class="sxs-lookup"><span data-stu-id="33e23-117">Solution</span></span>

<span data-ttu-id="33e23-118">Diese Lösung umfasst zwei Teile:</span><span class="sxs-lookup"><span data-stu-id="33e23-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="33e23-119">Ein Office-Skript zum Berechnen und Extrahieren Excel Diagramms und der Tabelle</span><span class="sxs-lookup"><span data-stu-id="33e23-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="33e23-120">Ein Power Automate Fluss, um das Skript aufzurufen und die Ergebnisse per E-Mail zu senden.</span><span class="sxs-lookup"><span data-stu-id="33e23-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="33e23-121">Ein Beispiel dazu finden Sie unter [Erstellen eines automatisierten Workflows mit Power Automate.](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)</span><span class="sxs-lookup"><span data-stu-id="33e23-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="33e23-122">Beispielcode: Berechnen und Extrahieren Excel Diagramms und der Tabelle</span><span class="sxs-lookup"><span data-stu-id="33e23-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="33e23-123">Das folgende Skript berechnet und extrahiert ein Excel Diagramm und eine Tabelle.</span><span class="sxs-lookup"><span data-stu-id="33e23-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="33e23-124">Laden Sie die Beispieldatei <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> herunter, und verwenden Sie sie mit diesem Skript, um sie selbst auszuprobieren!</span><span class="sxs-lookup"><span data-stu-id="33e23-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="33e23-125">Power Automate Fluss: Senden einer E-Mail an diagramm- und tabellenbilder</span><span class="sxs-lookup"><span data-stu-id="33e23-125">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="33e23-126">Dieser Fluss führt das Skript aus und sendet eine E-Mail an die zurückgegebenen Bilder.</span><span class="sxs-lookup"><span data-stu-id="33e23-126">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="33e23-127">Erstellen Sie einen neuen **Instant Cloud Flow.**</span><span class="sxs-lookup"><span data-stu-id="33e23-127">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="33e23-128">Wählen Sie **manuell einen Fluss auslösen** und erstellen drücken. </span><span class="sxs-lookup"><span data-stu-id="33e23-128">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="33e23-129">Fügen Sie einen **neuen Schritt** hinzu, der den connector Excel **Online (Business)** mit der **Skriptaktion ausführen** verwendet.</span><span class="sxs-lookup"><span data-stu-id="33e23-129">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="33e23-130">Verwenden Sie die folgenden Werte für die Aktion:</span><span class="sxs-lookup"><span data-stu-id="33e23-130">Use the following values for the action:</span></span>
    * <span data-ttu-id="33e23-131">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="33e23-131">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="33e23-132">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="33e23-132">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="33e23-133">**Datei:** Ihre Arbeitsmappe ([ausgewählt mit der Dateiauswahl](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="33e23-133">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="33e23-134">**Skript:** Ihr Skriptname</span><span class="sxs-lookup"><span data-stu-id="33e23-134">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Der fertige Excel Online (Business)-Connector in Power Automate.":::
1. <span data-ttu-id="33e23-136">In diesem Beispiel wird Outlook als E-Mail-Client verwendet.</span><span class="sxs-lookup"><span data-stu-id="33e23-136">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="33e23-137">Sie können einen beliebigen E-Mail-Connector verwenden, der Power Automate unterstützt. Bei den restlichen Schritten wird jedoch davon ausgegangen, dass Sie Outlook ausgewählt haben.</span><span class="sxs-lookup"><span data-stu-id="33e23-137">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="33e23-138">Fügen Sie einen **neuen Schritt** hinzu, der den **Office 365 Outlook** Connector und die Aktion **"Senden und E-Mail(V2)"** verwendet.</span><span class="sxs-lookup"><span data-stu-id="33e23-138">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="33e23-139">Verwenden Sie die folgenden Werte für die Aktion:</span><span class="sxs-lookup"><span data-stu-id="33e23-139">Use the following values for the action:</span></span>
    * <span data-ttu-id="33e23-140">**An:** Ihr Test-E-Mail-Konto (oder persönliche E-Mail)</span><span class="sxs-lookup"><span data-stu-id="33e23-140">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="33e23-141">**Betreff:** Bitte überprüfen Sie die Berichtsdaten</span><span class="sxs-lookup"><span data-stu-id="33e23-141">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="33e23-142">Wählen Sie für das Feld **"Text"** die Option "Codeansicht" ( `</>` ) aus, und geben Sie Folgendes ein:</span><span class="sxs-lookup"><span data-stu-id="33e23-142">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Der fertige Office 365 Outlook-Connector in Power Automate.":::
1. <span data-ttu-id="33e23-144">Speichern Sie den Flow, und testen Sie ihn.</span><span class="sxs-lookup"><span data-stu-id="33e23-144">Save the flow and try it out.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="33e23-145">Schulungsvideo: Extrahieren und E-Mail-Bilder von Diagrammen und Tabellen</span><span class="sxs-lookup"><span data-stu-id="33e23-145">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="33e23-146">[Sehen Sie sich dieses Beispiel auf YouTube an.](https://youtu.be/152GJyqc-Kw)</span><span class="sxs-lookup"><span data-stu-id="33e23-146">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
