---
title: Beispielskripts für Office-Skripts in Excel im Internet
description: Eine Sammlung von Codebeispielen, die mit Office-Skripts in Excel im Internet verwendet werden sollen.
ms.date: 06/18/2020
localization_priority: Normal
ms.openlocfilehash: bfa6679595e6e28cc5d2ae3e3e487fd3e77738aa
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878675"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="cdd06-103">Beispielskripts für Office-Skripts in Excel im Internet (Vorschau)</span><span class="sxs-lookup"><span data-stu-id="cdd06-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="cdd06-104">Die folgenden Beispiele sind einfache Skripts, mit denen Sie Ihre eigenen Arbeitsmappen ausprobieren können.</span><span class="sxs-lookup"><span data-stu-id="cdd06-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="cdd06-105">So verwenden Sie Sie in Excel im Internet:</span><span class="sxs-lookup"><span data-stu-id="cdd06-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="cdd06-106">Öffnen Sie die Registerkarte **Automatisieren**.</span><span class="sxs-lookup"><span data-stu-id="cdd06-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="cdd06-107">Drücken Sie **Code-Editor**.</span><span class="sxs-lookup"><span data-stu-id="cdd06-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="cdd06-108">Klicken Sie im Aufgabenbereich des Code-Editors auf **Neues Skript** .</span><span class="sxs-lookup"><span data-stu-id="cdd06-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="cdd06-109">Ersetzen Sie das gesamte Skript durch das Beispiel Ihrer Wahl.</span><span class="sxs-lookup"><span data-stu-id="cdd06-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="cdd06-110">Klicken Sie im Aufgabenbereich des Code-Editors auf **Ausführen** .</span><span class="sxs-lookup"><span data-stu-id="cdd06-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="cdd06-111">Grundlagen der Skripterstellung</span><span class="sxs-lookup"><span data-stu-id="cdd06-111">Scripting basics</span></span>

<span data-ttu-id="cdd06-112">In diesen Beispielen werden grundlegende Bausteine für Office-Skripts veranschaulicht.</span><span class="sxs-lookup"><span data-stu-id="cdd06-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="cdd06-113">Fügen Sie diese zu Ihren Skripts hinzu, um Ihre Lösung zu erweitern und häufige Probleme zu lösen.</span><span class="sxs-lookup"><span data-stu-id="cdd06-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="cdd06-114">Lesen und Protokollieren einer Zelle</span><span class="sxs-lookup"><span data-stu-id="cdd06-114">Read and log one cell</span></span>

<span data-ttu-id="cdd06-115">In diesem Beispiel wird der Wert von **a1** gelesen und in der Konsole gedruckt.</span><span class="sxs-lookup"><span data-stu-id="cdd06-115">This sample reads the value of **A1** and prints it to the console.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a><span data-ttu-id="cdd06-116">Lesen der aktiven Zelle</span><span class="sxs-lookup"><span data-stu-id="cdd06-116">Read the active cell</span></span>

<span data-ttu-id="cdd06-117">Dieses Skript protokolliert den Wert der aktuellen aktiven Zelle.</span><span class="sxs-lookup"><span data-stu-id="cdd06-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="cdd06-118">Wenn mehrere Zellen ausgewählt sind, wird die Zelle am weitesten links protokolliert.</span><span class="sxs-lookup"><span data-stu-id="cdd06-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="cdd06-119">Ändern einer benachbarten Zelle</span><span class="sxs-lookup"><span data-stu-id="cdd06-119">Change an adjacent cell</span></span>

<span data-ttu-id="cdd06-120">Dieses Skript ruft benachbarte Zellen mit relativen Verweisen ab.</span><span class="sxs-lookup"><span data-stu-id="cdd06-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="cdd06-121">Beachten Sie Folgendes: Wenn sich die aktive Zelle in der obersten Zeile befindet, schlägt ein Teil des Skripts fehl, da er auf die Zelle oberhalb der aktuell ausgewählten Zelle verweist.</span><span class="sxs-lookup"><span data-stu-id="cdd06-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently active cell in the workbook.
  let activeCell = workbook.getActiveCell();
  console.log(`The active cell's address is: ${activeCell.getAddress()}`);

  // Get the cell to the right of the active cell and set its value and color.
  let rightCell = activeCell.getOffsetRange(0,1);
  rightCell.setValue("Right cell");
  console.log(`The right cell's address is: ${rightCell.getAddress()}`);
  rightCell.getFormat().getFont().setColor("Magenta");
  rightCell.getFormat().getFill().setColor("Cyan");

  // Get the cell to the above of the active cell and set its value and color.
  // Note that this operation will fail if the active cell is in the top row.
  let aboveCell = activeCell.getOffsetRange(-1, 0);
  aboveCell.setValue("Above cell");
  console.log(`The above cell's address is: ${aboveCell.getAddress()}`);
  aboveCell.getFormat().getFont().setColor("White");
  aboveCell.getFormat().getFill().setColor("Black");
}
```

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="cdd06-122">Ändern aller benachbarten Zellen</span><span class="sxs-lookup"><span data-stu-id="cdd06-122">Change all adjacent cells</span></span>

<span data-ttu-id="cdd06-123">Dieses Skript kopiert die Formatierung der aktiven Zelle in die benachbarten Zellen.</span><span class="sxs-lookup"><span data-stu-id="cdd06-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="cdd06-124">Beachten Sie, dass dieses Skript nur funktioniert, wenn sich die aktive Zelle nicht auf einem Rand des Arbeitsblatts befindet.</span><span class="sxs-lookup"><span data-stu-id="cdd06-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  let activeCell = workbook.getActiveCell();

  // Get the cell that's one row above and one column to the left of the active cell.
  let cornerCell = activeCell.getOffsetRange(-1,-1);

  // Get a range that includes all the cells surrounding the active cell.
  let surroundingRange = cornerCell.getResizedRange(2, 2)

  // Copy the formatting from the active cell to the new range.
  surroundingRange.copyFrom(
    activeCell, /* The source range. */
    ExcelScript.RangeCopyType.formats /* What to copy. */
    );
}
```

### <a name="work-with-dates"></a><span data-ttu-id="cdd06-125">Arbeiten mit Datumsangaben</span><span class="sxs-lookup"><span data-stu-id="cdd06-125">Work with dates</span></span>

<span data-ttu-id="cdd06-126">In den Beispielen in diesem Abschnitt wird gezeigt, wie das JavaScript- [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) -Objekt verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="cdd06-126">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="cdd06-127">Im folgenden Beispiel werden das aktuelle Datum und die Uhrzeit abgerufen, und anschließend werden diese Werte in zwei Zellen im aktiven Arbeitsblatt geschrieben.</span><span class="sxs-lookup"><span data-stu-id="cdd06-127">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the cells at A1 and B1.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");
  let timeRange = workbook.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.setValue(date.toLocaleDateString());

  // Add the time string to B1.
  timeRange.setValue(date.toLocaleTimeString());
}
```

<span data-ttu-id="cdd06-128">Im nächsten Beispiel wird ein Datum gelesen, das in Excel gespeichert und in ein JavaScript-Date-Objekt übersetzt wird.</span><span class="sxs-lookup"><span data-stu-id="cdd06-128">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="cdd06-129">Es verwendet die [numerische Seriennummer des Datums](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) als Eingabe für das JavaScript-Datum.</span><span class="sxs-lookup"><span data-stu-id="cdd06-129">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue();
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a><span data-ttu-id="cdd06-130">Anzeigen von Daten</span><span class="sxs-lookup"><span data-stu-id="cdd06-130">Display data</span></span>

<span data-ttu-id="cdd06-131">In diesen Beispielen wird gezeigt, wie Sie mit Arbeitsblattdaten arbeiten und Benutzern eine bessere Ansicht oder Organisation bieten.</span><span class="sxs-lookup"><span data-stu-id="cdd06-131">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="cdd06-132">Anwenden bedingter Formatierung</span><span class="sxs-lookup"><span data-stu-id="cdd06-132">Apply conditional formatting</span></span>

<span data-ttu-id="cdd06-133">In diesem Beispiel wird die bedingte Formatierung auf den aktuell verwendeten Bereich im Arbeitsblatt angewendet.</span><span class="sxs-lookup"><span data-stu-id="cdd06-133">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="cdd06-134">Die bedingte Formatierung ist eine grüne Füllung für die oberen 10% der Werte.</span><span class="sxs-lookup"><span data-stu-id="cdd06-134">The conditional formatting is a green fill for the top 10% of values.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.addConditionalFormat(ExcelScript.ConditionalFormatType.topBottom)
  conditionalFormat.getTopBottom().getFormat().getFill().setColor("green");
  conditionalFormat.getTopBottom().setRule({
    rank: 10, // The percentage threshold.
    type: ExcelScript.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  });
}
```

### <a name="create-a-sorted-table"></a><span data-ttu-id="cdd06-135">Erstellen einer sortierten Tabelle</span><span class="sxs-lookup"><span data-stu-id="cdd06-135">Create a sorted table</span></span>

<span data-ttu-id="cdd06-136">In diesem Beispiel wird eine Tabelle aus dem verwendeten Bereich des aktuellen Arbeitsblatts erstellt und dann basierend auf der ersten Spalte sortiert.</span><span class="sxs-lookup"><span data-stu-id="cdd06-136">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.addTable(usedRange, true);

  // Sort the table using the first column.
  newTable.getSort().apply([{ key: 0, ascending: true }]);
}
```

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="cdd06-137">Protokollieren der "Gesamtsumme"-Werte aus einer PivotTable</span><span class="sxs-lookup"><span data-stu-id="cdd06-137">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="cdd06-138">In diesem Beispiel wird die erste PivotTable in der Arbeitsmappe gesucht und die Werte in den Zellen "Gesamtsumme" (wie in der Abbildung unten grün hervorgehoben) protokolliert.</span><span class="sxs-lookup"><span data-stu-id="cdd06-138">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

![Eine PivotTable mit Frucht Umsatz, wobei die Zeile "Gesamtergebnis" grün hervorgehoben ist.](../images/sample-pivottable-grand-total-row.png)

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getRangeBetweenHeaderAndTotal();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

## <a name="scenario-samples"></a><span data-ttu-id="cdd06-140">Szenario-Beispiele</span><span class="sxs-lookup"><span data-stu-id="cdd06-140">Scenario samples</span></span>

<span data-ttu-id="cdd06-141">Beispiele für größere Lösungen aus der realen Welt finden Sie unter [Beispielszenarien für Office-Skripts](scenarios/sample-scenario-overview.md).</span><span class="sxs-lookup"><span data-stu-id="cdd06-141">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="cdd06-142">Neue Beispiele vorschlagen</span><span class="sxs-lookup"><span data-stu-id="cdd06-142">Suggest new samples</span></span>

<span data-ttu-id="cdd06-143">Wir begrüßen Vorschläge für neue Beispiele.</span><span class="sxs-lookup"><span data-stu-id="cdd06-143">We welcome suggestions for new samples.</span></span> <span data-ttu-id="cdd06-144">Wenn es ein gängiges Szenario gibt, das anderen Skript Entwicklern helfen würde, teilen Sie uns dies im Abschnitt Feedback unten.</span><span class="sxs-lookup"><span data-stu-id="cdd06-144">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
