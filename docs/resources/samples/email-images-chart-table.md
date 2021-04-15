---
title: E-Mail der Bilder eines Excel-Diagramms und einer Tabelle
description: Erfahren Sie, wie Sie Office-Skripts und Power Automate verwenden, um die Bilder eines Excel-Diagramms und einer Tabelle zu extrahieren und per E-Mail zu senden.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: de3cf16537cb12db45d4d465d367d797d053afc4
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754810"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="18d54-103">Verwenden von Office-Skripts und Power Automate zum E-Mail-Senden von Bildern eines Diagramms und einer Tabelle</span><span class="sxs-lookup"><span data-stu-id="18d54-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="18d54-104">In diesem Beispiel werden Office-Skripts und Power Automate zum Erstellen eines Diagramms verwendet.</span><span class="sxs-lookup"><span data-stu-id="18d54-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="18d54-105">Anschließend werden Bilder des Diagramms und seiner Basistabelle per E-Mail gesendet.</span><span class="sxs-lookup"><span data-stu-id="18d54-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="18d54-106">Beispielszenario</span><span class="sxs-lookup"><span data-stu-id="18d54-106">Example scenario</span></span>

* <span data-ttu-id="18d54-107">Berechnen Sie, um die neuesten Ergebnisse zu erhalten.</span><span class="sxs-lookup"><span data-stu-id="18d54-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="18d54-108">Diagramm erstellen.</span><span class="sxs-lookup"><span data-stu-id="18d54-108">Create chart.</span></span>
* <span data-ttu-id="18d54-109">Get chart and table images.</span><span class="sxs-lookup"><span data-stu-id="18d54-109">Get chart and table images.</span></span>
* <span data-ttu-id="18d54-110">Senden Sie eine E-Mail an die Bilder mit Power Automate.</span><span class="sxs-lookup"><span data-stu-id="18d54-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="18d54-111">_Eingabedaten_</span><span class="sxs-lookup"><span data-stu-id="18d54-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Ein Arbeitsblatt mit einer Tabelle mit Eingabedaten.":::

<span data-ttu-id="18d54-113">_Ausgabediagramm_</span><span class="sxs-lookup"><span data-stu-id="18d54-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="Das vom Kunden fällige Spaltendiagramm.":::

<span data-ttu-id="18d54-115">_E-Mails, die über den Power Automate-Fluss empfangen wurden_</span><span class="sxs-lookup"><span data-stu-id="18d54-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="Die vom Fluss gesendete E-Mail, die das im Textkörper eingebettete Excel-Diagramm zeigt.":::

## <a name="solution"></a><span data-ttu-id="18d54-117">Lösung</span><span class="sxs-lookup"><span data-stu-id="18d54-117">Solution</span></span>

<span data-ttu-id="18d54-118">Diese Lösung besteht aus zwei Teilen:</span><span class="sxs-lookup"><span data-stu-id="18d54-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="18d54-119">Ein Office-Skript zum Berechnen und Extrahieren von Excel-Diagrammen und -Tabellen</span><span class="sxs-lookup"><span data-stu-id="18d54-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="18d54-120">Ein Power Automate-Fluss zum Aufrufen des Skripts und zum Senden der Ergebnisse per E-Mail.</span><span class="sxs-lookup"><span data-stu-id="18d54-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="18d54-121">Ein Beispiel dazu finden Sie unter Erstellen eines automatisierten [Workflows mit Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="18d54-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="18d54-122">Beispielcode: Berechnen und Extrahieren von Excel-Diagrammen und -Tabellen</span><span class="sxs-lookup"><span data-stu-id="18d54-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="18d54-123">Das folgende Skript berechnet und extrahiert ein Excel-Diagramm und eine Tabelle.</span><span class="sxs-lookup"><span data-stu-id="18d54-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="18d54-124">Laden Sie die Beispieldatei <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> und verwenden Sie sie mit diesem Skript, um sie selbst auszuprobieren!</span><span class="sxs-lookup"><span data-stu-id="18d54-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');
  const targetRange = updateRange(chartSheet, selectColumns);

  // Insert chart on sheet 'Sheet1'.
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="18d54-125">Schulungsvideo: Extrahieren und E-Mail-Bilder von Diagramm und Tabelle</span><span class="sxs-lookup"><span data-stu-id="18d54-125">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="18d54-126">[![Sehen Sie sich schritt-für-Schritt-Video zum Extrahieren und E-Mail-Bilder von Diagramm und Tabelle an](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Schritt-für-Schritt-Video zum Extrahieren und E-Mail-Bilder von Diagramm und Tabelle")</span><span class="sxs-lookup"><span data-stu-id="18d54-126">[![Watch step-by-step video on how to extract and email images of chart and table](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Step-by-step video on how to extract and email images of chart and table")</span></span>
