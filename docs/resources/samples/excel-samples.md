---
title: Grundlegende Skripts für Office Skripts in Excel im Web
description: Eine Sammlung von Codebeispielen, die mit Office Skripts in Excel im Web verwendet werden sollen.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 3aaaa7fe8769f6dcd658ae91c577956b56033051
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313939"
---
# <a name="basic-scripts-for-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="7ab53-103">Grundlegende Skripts für Office Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="7ab53-103">Basic scripts for Office Scripts in Excel on the web</span></span>

<span data-ttu-id="7ab53-104">Die folgenden Beispiele sind einfache Skripts, mit denen Sie Ihre eigenen Arbeitsmappen ausprobieren können.</span><span class="sxs-lookup"><span data-stu-id="7ab53-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="7ab53-105">So verwenden Sie sie in Excel im Web:</span><span class="sxs-lookup"><span data-stu-id="7ab53-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="7ab53-106">Öffnen Sie die Registerkarte **Automatisieren**.</span><span class="sxs-lookup"><span data-stu-id="7ab53-106">Open the **Automate** tab.</span></span>
1. <span data-ttu-id="7ab53-107">Wählen Sie **New Script** aus.</span><span class="sxs-lookup"><span data-stu-id="7ab53-107">Select **New Script**.</span></span>
1. <span data-ttu-id="7ab53-108">Ersetzen Sie das gesamte Skript durch das Beispiel Ihrer Wahl.</span><span class="sxs-lookup"><span data-stu-id="7ab53-108">Replace the entire script with the sample of your choice.</span></span>
1. <span data-ttu-id="7ab53-109">Wählen Sie im Aufgabenbereich des Code-Editors **"Ausführen"** aus.</span><span class="sxs-lookup"><span data-stu-id="7ab53-109">Select **Run** in the Code Editor's task pane.</span></span>

## <a name="script-basics"></a><span data-ttu-id="7ab53-110">Skriptgrundlagen</span><span class="sxs-lookup"><span data-stu-id="7ab53-110">Script basics</span></span>

<span data-ttu-id="7ab53-111">Diese Beispiele veranschaulichen grundlegende Bausteine für Office Skripts.</span><span class="sxs-lookup"><span data-stu-id="7ab53-111">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="7ab53-112">Erweitern Sie diese Skripts, um Ihre Lösung zu erweitern und allgemeine Probleme zu lösen.</span><span class="sxs-lookup"><span data-stu-id="7ab53-112">Expand these scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="7ab53-113">Lesen und Protokollieren einer Zelle</span><span class="sxs-lookup"><span data-stu-id="7ab53-113">Read and log one cell</span></span>

<span data-ttu-id="7ab53-114">In diesem Beispiel wird der Wert von **A1** gelesen und in der Konsole gedruckt.</span><span class="sxs-lookup"><span data-stu-id="7ab53-114">This sample reads the value of **A1** and prints it to the console.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a><span data-ttu-id="7ab53-115">Lesen der aktiven Zelle</span><span class="sxs-lookup"><span data-stu-id="7ab53-115">Read the active cell</span></span>

<span data-ttu-id="7ab53-116">Dieses Skript protokolliert den Wert der aktuellen aktiven Zelle.</span><span class="sxs-lookup"><span data-stu-id="7ab53-116">This script logs the value of the current active cell.</span></span> <span data-ttu-id="7ab53-117">Wenn mehrere Zellen ausgewählt sind, wird die oberste linke Zelle protokolliert.</span><span class="sxs-lookup"><span data-stu-id="7ab53-117">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="7ab53-118">Ändern einer angrenzenden Zelle</span><span class="sxs-lookup"><span data-stu-id="7ab53-118">Change an adjacent cell</span></span>

<span data-ttu-id="7ab53-119">Dieses Skript ruft benachbarte Zellen mit relativen Verweisen ab.</span><span class="sxs-lookup"><span data-stu-id="7ab53-119">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="7ab53-120">Beachten Sie, dass ein Teil des Skripts fehlschlägt, wenn sich die aktive Zelle in der oberen Zeile befindet, da sie auf die Zelle über der aktuell ausgewählten Zelle verweist.</span><span class="sxs-lookup"><span data-stu-id="7ab53-120">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

```TypeScript
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

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="7ab53-121">Ändern aller angrenzenden Zellen</span><span class="sxs-lookup"><span data-stu-id="7ab53-121">Change all adjacent cells</span></span>

<span data-ttu-id="7ab53-122">Dieses Skript kopiert die Formatierung in der aktiven Zelle in die benachbarten Zellen.</span><span class="sxs-lookup"><span data-stu-id="7ab53-122">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="7ab53-123">Beachten Sie, dass dieses Skript nur funktioniert, wenn sich die aktive Zelle nicht an einem Rand des Arbeitsblatts befindet.</span><span class="sxs-lookup"><span data-stu-id="7ab53-123">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

```TypeScript
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

### <a name="change-each-individual-cell-in-a-range"></a><span data-ttu-id="7ab53-124">Ändern jeder einzelnen Zelle in einem Bereich</span><span class="sxs-lookup"><span data-stu-id="7ab53-124">Change each individual cell in a range</span></span>

<span data-ttu-id="7ab53-125">Dieses Skript durchläuft den aktuell ausgewählten Bereich in einer Schleife.</span><span class="sxs-lookup"><span data-stu-id="7ab53-125">This script loops over the currently select range.</span></span> <span data-ttu-id="7ab53-126">Es löscht die aktuelle Formatierung und legt die Füllfarbe in jeder Zelle auf eine zufällige Farbe fest.</span><span class="sxs-lookup"><span data-stu-id="7ab53-126">It clears the current formatting and sets the fill color in each cell to a random color.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected range.
  let range = workbook.getSelectedRange();

  // Get the size boundaries of the range.
  let rows = range.getRowCount();
  let cols = range.getColumnCount();

  // Clear any existing formatting
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Iterate over the range.
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      // Generate a random color hex-code.
      let colorString = `#${Math.random().toString(16).substr(-6)}`;

      // Set the color of the current cell to that random hex-code.
      range.getCell(row, col).getFormat().getFill().setColor(colorString);
    }
  }
}
```

### <a name="get-groups-of-cells-based-on-special-criteria"></a><span data-ttu-id="7ab53-127">Abrufen von Zellgruppen basierend auf speziellen Kriterien</span><span class="sxs-lookup"><span data-stu-id="7ab53-127">Get groups of cells based on special criteria</span></span>

<span data-ttu-id="7ab53-128">Dieses Skript ruft alle leeren Zellen im verwendeten Bereich des aktuellen Arbeitsblatts ab.</span><span class="sxs-lookup"><span data-stu-id="7ab53-128">This script gets all the blank cells in the current worksheet's used range.</span></span> <span data-ttu-id="7ab53-129">Anschließend werden alle Zellen mit gelbem Hintergrund hervorgehoben.</span><span class="sxs-lookup"><span data-stu-id="7ab53-129">It then highlights all those cells with a yellow background.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the current used range.
    let range = workbook.getActiveWorksheet().getUsedRange();
    
    // Get all the blank cells.
    let blankCells = range.getSpecialCells(ExcelScript.SpecialCellType.blanks);

    // Highlight the blank cells with a yellow background.
    blankCells.getFormat().getFill().setColor("yellow");
}
```

## <a name="collections"></a><span data-ttu-id="7ab53-130">Auflistungen</span><span class="sxs-lookup"><span data-stu-id="7ab53-130">Collections</span></span>

<span data-ttu-id="7ab53-131">Diese Beispiele funktionieren mit Auflistungen von Objekten in der Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="7ab53-131">These samples work with collections of objects in the workbook.</span></span>

### <a name="iterate-over-collections"></a><span data-ttu-id="7ab53-132">Durchlaufen von Sammlungen</span><span class="sxs-lookup"><span data-stu-id="7ab53-132">Iterate over collections</span></span>

<span data-ttu-id="7ab53-133">Dieses Skript ruft die Namen aller Arbeitsblätter in der Arbeitsmappe ab und protokolliert sie.</span><span class="sxs-lookup"><span data-stu-id="7ab53-133">This script gets and logs the names of all the worksheets in the workbook.</span></span> <span data-ttu-id="7ab53-134">Außerdem werden die Registerkartenfarben auf eine zufällige Farbe festgelegt.</span><span class="sxs-lookup"><span data-stu-id="7ab53-134">It also sets the their tab colors to a random color.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get all the worksheets in the workbook.
  let sheets = workbook.getWorksheets();

  // Get a list of all the worksheet names.
  let names = sheets.map ((sheet) => sheet.getName());

  // Write in the console all the worksheet names and the total count.
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  
  // Set the tab color each worksheet to a random color
  for (let sheet of sheets) {
    // Generate a random color hex-code.
    let colorString = `#${Math.random().toString(16).substr(-6)}`;

    // Set the color of the current worksheet's tab to that random hex-code.
    sheet.setTabColor(colorString);
  }
}
```

### <a name="query-and-delete-from-a-collection"></a><span data-ttu-id="7ab53-135">Abfragen und Löschen aus einer Auflistung</span><span class="sxs-lookup"><span data-stu-id="7ab53-135">Query and delete from a collection</span></span>

<span data-ttu-id="7ab53-136">Dieses Skript erstellt ein neues Arbeitsblatt.</span><span class="sxs-lookup"><span data-stu-id="7ab53-136">This script creates a new worksheet.</span></span> <span data-ttu-id="7ab53-137">Es wird nach einer vorhandenen Kopie des Arbeitsblatts gesucht und vor dem Erstellen eines neuen Blatts gelöscht.</span><span class="sxs-lookup"><span data-stu-id="7ab53-137">It checks for an existing copy of the worksheet and deletes it before making a new sheet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added.
  let name = "Index";

  // Get any worksheet with that name.
  let sheet = workbook.getWorksheet("Index");
  
  // If `null` wasn't returned, then there's already a worksheet with that name.
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Delete the sheet.
    sheet.delete();
  }
  
  // Add a blank worksheet with the name "Index".
  // Note that this code runs regardless of whether an existing sheet was deleted.
  console.log(`Adding the worksheet named ${name}.`);
  let newSheet = workbook.addWorksheet("Index");

  // Switch to the new worksheet.
  newSheet.activate();
}
```

## <a name="dates"></a><span data-ttu-id="7ab53-138">Datumsangaben</span><span class="sxs-lookup"><span data-stu-id="7ab53-138">Dates</span></span>

<span data-ttu-id="7ab53-139">Die Beispiele in diesem Abschnitt zeigen, wie Sie das JavaScript [Date-Objekt](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) verwenden.</span><span class="sxs-lookup"><span data-stu-id="7ab53-139">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="7ab53-140">Das folgende Beispiel ruft das aktuelle Datum und die aktuelle Uhrzeit ab und schreibt diese Werte dann in zwei Zellen im aktiven Arbeitsblatt.</span><span class="sxs-lookup"><span data-stu-id="7ab53-140">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="7ab53-141">Im nächsten Beispiel wird ein Datum gelesen, das in Excel gespeichert ist, und in ein JavaScript Date-Objekt übersetzt.</span><span class="sxs-lookup"><span data-stu-id="7ab53-141">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="7ab53-142">Es verwendet die [numerische Seriennummer des Datums](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) als Eingabe für das JavaScript-Datum.</span><span class="sxs-lookup"><span data-stu-id="7ab53-142">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue() as number;
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a><span data-ttu-id="7ab53-143">Anzeigen von Daten</span><span class="sxs-lookup"><span data-stu-id="7ab53-143">Display data</span></span>

<span data-ttu-id="7ab53-144">Diese Beispiele veranschaulichen, wie Sie mit Arbeitsblattdaten arbeiten und Benutzern eine bessere Ansicht oder Organisation bieten.</span><span class="sxs-lookup"><span data-stu-id="7ab53-144">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="7ab53-145">Anwenden bedingter Formatierung</span><span class="sxs-lookup"><span data-stu-id="7ab53-145">Apply conditional formatting</span></span>

<span data-ttu-id="7ab53-146">In diesem Beispiel wird dem aktuell verwendeten Bereich im Arbeitsblatt eine bedingte Formatierung zugewiesen.</span><span class="sxs-lookup"><span data-stu-id="7ab53-146">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="7ab53-147">Die bedingte Formatierung ist eine grüne Füllung für die oberen 10 % der Werte.</span><span class="sxs-lookup"><span data-stu-id="7ab53-147">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="7ab53-148">Erstellen einer sortierten Tabelle</span><span class="sxs-lookup"><span data-stu-id="7ab53-148">Create a sorted table</span></span>

<span data-ttu-id="7ab53-149">In diesem Beispiel wird eine Tabelle aus dem verwendeten Bereich des aktuellen Arbeitsblatts erstellt und anschließend basierend auf der ersten Spalte sortiert.</span><span class="sxs-lookup"><span data-stu-id="7ab53-149">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="7ab53-150">Protokollieren der "Gesamtsumme"-Werte aus einer PivotTable</span><span class="sxs-lookup"><span data-stu-id="7ab53-150">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="7ab53-151">In diesem Beispiel wird die erste PivotTable in der Arbeitsmappe gesucht und die Werte in den Zellen Gesamtsumme protokolliert (wie in der abbildung unten grün hervorgehoben).</span><span class="sxs-lookup"><span data-stu-id="7ab53-151">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="Eine PivotTable mit Obstumsätzen, wobei die Ergebniszeile grün hervorgehoben ist.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getBodyAndTotalRange();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

### <a name="create-a-drop-down-list-using-data-validation"></a><span data-ttu-id="7ab53-153">Erstellen einer Dropdownliste mithilfe der Datenüberprüfung</span><span class="sxs-lookup"><span data-stu-id="7ab53-153">Create a drop-down list using data validation</span></span>

<span data-ttu-id="7ab53-154">Dieses Skript erstellt eine Dropdown-Auswahlliste für eine Zelle.</span><span class="sxs-lookup"><span data-stu-id="7ab53-154">This script creates a drop-down selection list for a cell.</span></span> <span data-ttu-id="7ab53-155">Es verwendet die vorhandenen Werte des ausgewählten Bereichs als Auswahl für die Liste.</span><span class="sxs-lookup"><span data-stu-id="7ab53-155">It uses the existing values of the selected range as the choices for the list.</span></span>

:::image type="content" source="../../images/sample-data-validation.png" alt-text="Ein Arbeitsblatt mit einem Bereich von drei Zellen, der die Farbauswahl &quot;Rot, Blau, Grün&quot; und daneben die gleichen Auswahlmöglichkeiten enthält, die in einer Dropdownliste angezeigt werden.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the values for data validation.
  let selectedRange = workbook.getSelectedRange();
  let rangeValues = selectedRange.getValues();

  // Convert the values into a comma-delimited string.
  let dataValidationListString = "";
  rangeValues.forEach((rangeValueRow) => {
    rangeValueRow.forEach((value) => {
      dataValidationListString += value + ",";
    });
  });

  // Clear the old range.
  selectedRange.clear(ExcelScript.ClearApplyTo.contents);

  // Apply the data validation to the first cell in the selected range.
  let targetCell = selectedRange.getCell(0,0);
  let dataValidation = targetCell.getDataValidation();

  // Set the content of the drop-down list.
  dataValidation.setRule({
      list: {
        inCellDropDown: true,
        source: dataValidationListString
      }
    });
}
```

## <a name="formulas"></a><span data-ttu-id="7ab53-157">Formeln</span><span class="sxs-lookup"><span data-stu-id="7ab53-157">Formulas</span></span>

<span data-ttu-id="7ab53-158">Diese Beispiele verwenden Excel Formeln und zeigen, wie Sie mit ihnen in Skripts arbeiten.</span><span class="sxs-lookup"><span data-stu-id="7ab53-158">These samples use Excel formulas and show how to work with them in scripts.</span></span>

### <a name="single-formula"></a><span data-ttu-id="7ab53-159">Einzelne Formel</span><span class="sxs-lookup"><span data-stu-id="7ab53-159">Single formula</span></span>

<span data-ttu-id="7ab53-160">Dieses Skript legt die Formel einer Zelle fest und zeigt dann an, wie Excel die Formel und den Wert der Zelle separat speichert.</span><span class="sxs-lookup"><span data-stu-id="7ab53-160">This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  // Set A1 to 2.
  let a1 = selectedSheet.getRange("A1");
  a1.setValue(2);

  // Set B1 to the formula =(2*A1), which should equal 4.
  let b1 = selectedSheet.getRange("B1")
  b1.setFormula("=(2*A1)");

  // Log the current results for `getFormula` and `getValue` at B1.
  console.log(`B1 - Formula: ${b1.getFormula()} | Value: ${b1.getValue()}`);
}
```

### <a name="handle-a-spill-error-returned-from-a-formula"></a><span data-ttu-id="7ab53-161">Behandeln eines `#SPILL!` von einer Formel zurückgegebenen Fehlers</span><span class="sxs-lookup"><span data-stu-id="7ab53-161">Handle a `#SPILL!` error returned from a formula</span></span>

<span data-ttu-id="7ab53-162">Dieses Skript transponiert den Bereich "A1:D2" mithilfe der TRANSPOSE-Funktion in "A4:B7".</span><span class="sxs-lookup"><span data-stu-id="7ab53-162">This script transposes the range "A1:D2" to "A4:B7" by using the TRANSPOSE function.</span></span> <span data-ttu-id="7ab53-163">Wenn die Transponieren zu einem `#SPILL` Fehler führt, löscht sie den Zielbereich und wendet die Formel erneut an.</span><span class="sxs-lookup"><span data-stu-id="7ab53-163">If the transpose results in a `#SPILL` error, it clears the target range and applies the formula again.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  // Use the data in A1:D2 for the sample.
  let dataAddress = "A1:D2"
  let inputRange = sheet.getRange(dataAddress);

  // Place the transposed data starting at A4.
  let targetStartCell = sheet.getRange("A4");

  // Compute the target range.
  let targetRange = targetStartCell.getResizedRange(inputRange.getColumnCount() - 1, inputRange.getRowCount() - 1);

  // Call the transpose helper function.
  targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);

  // Check if the range update resulted in a spill error.
  let checkValue = targetStartCell.getValue() as string;
  if (checkValue === '#SPILL!') {
    // Clear the target range and call the transpose function again.
    console.log("Target range has data that is preventing update. Clearing target range.");
    targetRange.clear();
    targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);
  }

  // Select the transposed range to highlight it.
  targetRange.select();
}
```

## <a name="suggest-new-samples"></a><span data-ttu-id="7ab53-164">Vorschlagen neuer Beispiele</span><span class="sxs-lookup"><span data-stu-id="7ab53-164">Suggest new samples</span></span>

<span data-ttu-id="7ab53-165">Wir freuen uns über Vorschläge für neue Beispiele.</span><span class="sxs-lookup"><span data-stu-id="7ab53-165">We welcome suggestions for new samples.</span></span> <span data-ttu-id="7ab53-166">Wenn es ein häufiges Szenario gibt, das anderen Skriptentwicklern helfen würde, teilen Sie uns dies bitte im Feedbackabschnitt am unteren Rand der Seite mit.</span><span class="sxs-lookup"><span data-stu-id="7ab53-166">If there is a common scenario that would help other script developers, please tell us in the feedback section at the bottom of the page.</span></span>

## <a name="see-also"></a><span data-ttu-id="7ab53-167">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="7ab53-167">See also</span></span>

* [<span data-ttu-id="7ab53-168">Adrianhi Ramamurthys "Range basics" auf YouTube</span><span class="sxs-lookup"><span data-stu-id="7ab53-168">Sudhi Ramamurthy's "Range basics" on YouTube</span></span>](https://youtu.be/4emjkOFdLBA)
* [<span data-ttu-id="7ab53-169">Office Skriptbeispiele und -szenarien</span><span class="sxs-lookup"><span data-stu-id="7ab53-169">Office Scripts samples and scenarios</span></span>](samples-overview.md)
* [<span data-ttu-id="7ab53-170">Aufzeichnen, Bearbeiten und Erstellen von Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="7ab53-170">Record, edit, and create Office Scripts in Excel on the web</span></span>](../../tutorials/excel-tutorial.md)
