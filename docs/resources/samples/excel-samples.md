---
title: Grundlegende Skripts für Office Skripts in Excel
description: Eine Sammlung von Codebeispielen für die Verwendung mit Office Skripts in Excel.
ms.date: 03/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8e28026b7a3498d477cce8b6dc5940da33a30f53
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393656"
---
# <a name="basic-scripts-for-office-scripts-in-excel"></a>Grundlegende Skripts für Office Skripts in Excel

Die folgenden Beispiele sind einfache Skripts, mit denen Sie Ihre eigenen Arbeitsmappen ausprobieren können. So verwenden Sie sie in Excel:

1. Öffnen Sie eine Arbeitsmappe in Excel im Web.
1. Öffnen Sie die Registerkarte **Automatisieren**.
1. Wählen Sie **Neues Skript** aus.
1. Ersetzen Sie das gesamte Skript durch das Beispiel Ihrer Wahl.
1. Wählen Sie im Aufgabenbereich des Code-Editors " **Ausführen** " aus.

## <a name="script-basics"></a>Skriptgrundlagen

Diese Beispiele veranschaulichen grundlegende Bausteine für Office Skripts. Erweitern Sie diese Skripts, um Ihre Lösung zu erweitern und häufige Probleme zu lösen.

### <a name="read-and-log-one-cell"></a>Lesen und Protokollieren einer Zelle

In diesem Beispiel wird der Wert von **A1** gelesen und in der Konsole gedruckt.

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

### <a name="read-the-active-cell"></a>Lesen der aktiven Zelle

Dieses Skript protokolliert den Wert der aktuellen aktiven Zelle. Wenn mehrere Zellen ausgewählt sind, wird die Zelle ganz oben links protokolliert.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a>Ändern einer angrenzenden Zelle

Dieses Skript ruft benachbarte Zellen mithilfe relativer Bezüge ab. Beachten Sie, dass ein Teil des Skripts fehlschlägt, wenn sich die aktive Zelle in der obersten Zeile befindet, da sie auf die Zelle oberhalb der aktuell ausgewählten Zelle verweist.

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

### <a name="change-all-adjacent-cells"></a>Ändern aller angrenzenden Zellen

Dieses Skript kopiert die Formatierung in der aktiven Zelle in die benachbarten Zellen. Beachten Sie, dass dieses Skript nur funktioniert, wenn sich die aktive Zelle nicht an einem Rand des Arbeitsblatts befindet.

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

### <a name="change-each-individual-cell-in-a-range"></a>Ändern jeder einzelnen Zelle in einem Bereich

Dieses Skript überläuft den aktuell ausgewählten Bereich in einer Schleife. Sie löscht die aktuelle Formatierung und legt die Füllfarbe in jeder Zelle auf eine zufällige Farbe fest.

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

### <a name="get-groups-of-cells-based-on-special-criteria"></a>Abrufen von Zellgruppen basierend auf speziellen Kriterien

Dieses Skript ruft alle leeren Zellen im verwendeten Bereich des aktuellen Arbeitsblatts ab. Anschließend werden alle Zellen mit einem gelben Hintergrund hervorgehoben.

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

## <a name="collections"></a>Sammlungen

Diese Beispiele funktionieren mit Auflistungen von Objekten in der Arbeitsmappe.

### <a name="iterate-over-collections"></a>Durchlaufen von Sammlungen

Dieses Skript ruft die Namen aller Arbeitsblätter in der Arbeitsmappe ab und protokolliert sie. Außerdem werden die Registerkartenfarben auf eine zufällige Farbe festgelegt.

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

### <a name="query-and-delete-from-a-collection"></a>Abfragen und Löschen aus einer Sammlung

Dieses Skript erstellt ein neues Arbeitsblatt. Sie sucht nach einer vorhandenen Kopie des Arbeitsblatts und löscht sie, bevor ein neues Blatt erstellt wird.

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

## <a name="dates"></a>Datumsangaben

Die Beispiele in diesem Abschnitt zeigen die Verwendung des JavaScript [Date-Objekts](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) .

Das folgende Beispiel ruft das aktuelle Datum und die aktuelle Uhrzeit ab und schreibt diese Werte dann in zwei Zellen im aktiven Arbeitsblatt.

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

Im nächsten Beispiel wird ein Datum gelesen, das in Excel gespeichert ist, und es in ein JavaScript Date-Objekt übersetzt. Es verwendet die numerische Seriennummer des Datums als Eingabe für das JavaScript-Datum. Diese Seriennummer wird im Artikel zur [FUNKTION NOW()](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) beschrieben.

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

## <a name="display-data"></a>Anzeigen von Daten

Diese Beispiele veranschaulichen, wie Sie mit Arbeitsblattdaten arbeiten und Benutzern eine bessere Ansicht oder Organisation bieten.

### <a name="apply-conditional-formatting"></a>Anwenden bedingter Formatierung

In diesem Beispiel wird bedingte Formatierung auf den aktuell verwendeten Bereich im Arbeitsblatt angewendet. Die bedingte Formatierung ist eine grüne Füllung für die obersten 10 % der Werte.

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

### <a name="create-a-sorted-table"></a>Erstellen einer sortierten Tabelle

In diesem Beispiel wird eine Tabelle aus dem verwendeten Bereich des aktuellen Arbeitsblatts erstellt und dann basierend auf der ersten Spalte sortiert.

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a>Protokollieren der "Gesamtsumme"-Werte aus einer PivotTable

In diesem Beispiel wird die erste PivotTable in der Arbeitsmappe gesucht und die Werte in den Zellen "Gesamtsumme" protokolliert (wie in der abbildung unten grün hervorgehoben).

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="Eine PivotTable mit Obstumsätzen, in der die Zeile &quot;Gesamtergebnis&quot; grün hervorgehoben ist.":::

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

### <a name="create-a-drop-down-list-using-data-validation"></a>Erstellen einer Dropdownliste mithilfe der Datenüberprüfung

Dieses Skript erstellt eine Dropdownauswahlliste für eine Zelle. Es verwendet die vorhandenen Werte des ausgewählten Bereichs als Auswahlmöglichkeiten für die Liste.

:::image type="content" source="../../images/sample-data-validation.png" alt-text="Ein Arbeitsblatt mit einem Bereich von drei Zellen mit den Farboptionen &quot;rot, blau, grün&quot; und daneben die gleichen Optionen, die in einer Dropdownliste angezeigt werden.":::

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

## <a name="formulas"></a>Formeln

Diese Beispiele verwenden Excel Formeln und zeigen, wie Sie mit ihnen in Skripts arbeiten.

### <a name="single-formula"></a>Einzelne Formel

Dieses Skript legt die Formel einer Zelle fest und zeigt dann an, wie Excel die Formel und den Wert der Zelle separat speichert.

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

### <a name="handle-a-spill-error-returned-from-a-formula"></a>Behandeln eines von einer `#SPILL!` Formel zurückgegebenen Fehlers

Dieses Skript transponiert den Bereich "A1:D2" mithilfe der MTRANS-Funktion in "A4:B7". Wenn die Transponierung zu einem `#SPILL` Fehler führt, wird der Zielbereich gelöscht und die Formel erneut angewendet.

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

## <a name="suggest-new-samples"></a>Vorschläge für neue Beispiele

Wir freuen uns über Vorschläge für neue Beispiele. Wenn es ein häufiges Szenario gibt, das anderen Skriptentwicklern helfen würde, teilen Sie uns dies bitte im Feedbackabschnitt am unteren Rand der Seite mit.

## <a name="see-also"></a>Siehe auch

* [Sudhi Ramamurthys "Range basics" auf YouTube](https://youtu.be/4emjkOFdLBA)
* [beispiele und Szenarien für Office Skripts](samples-overview.md)
* [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripts in Excel im Web](../../tutorials/excel-tutorial.md)
