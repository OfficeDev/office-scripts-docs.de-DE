---
title: Beispielskripts für Office-Skripts in Excel im Internet
description: Eine Sammlung von Codebeispielen, die mit Office-Skripts in Excel im Internet verwendet werden sollen.
ms.date: 07/16/2020
localization_priority: Normal
ms.openlocfilehash: fa330bfa284799e26ee2cf49800102072d66612b
ms.sourcegitcommit: 8d549884e68170f808d3d417104a4451a37da83c
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/22/2020
ms.locfileid: "45229603"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a>Beispielskripts für Office-Skripts in Excel im Internet (Vorschau)

Die folgenden Beispiele sind einfache Skripts, mit denen Sie Ihre eigenen Arbeitsmappen ausprobieren können. So verwenden Sie Sie in Excel im Internet:

1. Öffnen Sie die Registerkarte **Automatisieren**.
2. Drücken Sie **Code-Editor**.
3. Klicken Sie im Aufgabenbereich des Code-Editors auf **Neues Skript** .
4. Ersetzen Sie das gesamte Skript durch das Beispiel Ihrer Wahl.
5. Klicken Sie im Aufgabenbereich des Code-Editors auf **Ausführen** .

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a>Grundlagen der Skripterstellung

In diesen Beispielen werden grundlegende Bausteine für Office-Skripts veranschaulicht. Fügen Sie diese zu Ihren Skripts hinzu, um Ihre Lösung zu erweitern und häufige Probleme zu lösen.

### <a name="read-and-log-one-cell"></a>Lesen und Protokollieren einer Zelle

In diesem Beispiel wird der Wert von **a1** gelesen und in der Konsole gedruckt.

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

### <a name="read-the-active-cell"></a>Lesen der aktiven Zelle

Dieses Skript protokolliert den Wert der aktuellen aktiven Zelle. Wenn mehrere Zellen ausgewählt sind, wird die Zelle am weitesten links protokolliert.

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a>Ändern einer benachbarten Zelle

Dieses Skript ruft benachbarte Zellen mit relativen Verweisen ab. Beachten Sie Folgendes: Wenn sich die aktive Zelle in der obersten Zeile befindet, schlägt ein Teil des Skripts fehl, da er auf die Zelle oberhalb der aktuell ausgewählten Zelle verweist.

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

### <a name="change-all-adjacent-cells"></a>Ändern aller benachbarten Zellen

Dieses Skript kopiert die Formatierung der aktiven Zelle in die benachbarten Zellen. Beachten Sie, dass dieses Skript nur funktioniert, wenn sich die aktive Zelle nicht auf einem Rand des Arbeitsblatts befindet.

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

### <a name="change-each-individual-cell-in-a-range"></a>Ändern der einzelnen Zellen in einem Bereich

Dieses Skript durchläuft den aktuell ausgewählten Bereich. Es löscht die aktuelle Formatierung und legt die Füllfarbe in jeder Zelle auf eine zufällige Farbe fest.

```typescript
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

## <a name="collections"></a>Auflistungen

Diese Beispiele funktionieren mit Auflistungen von Objekten in der Arbeitsmappe.

### <a name="iterating-over-collections"></a>Durchlaufen von Auflistungen

Dieses Skript ruft die Namen aller Arbeitsblätter in der Arbeitsmappe ab und protokolliert sie. Außerdem werden die Farben ihrer Registerkarten auf eine zufällige Farbe festgelegt.

```typescript
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

## <a name="dates"></a>Datumsangaben

In den Beispielen in diesem Abschnitt wird gezeigt, wie das JavaScript- [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) -Objekt verwendet wird.

Im folgenden Beispiel werden das aktuelle Datum und die Uhrzeit abgerufen, und anschließend werden diese Werte in zwei Zellen im aktiven Arbeitsblatt geschrieben.

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

Im nächsten Beispiel wird ein Datum gelesen, das in Excel gespeichert und in ein JavaScript-Date-Objekt übersetzt wird. Es verwendet die [numerische Seriennummer des Datums](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) als Eingabe für das JavaScript-Datum.

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

## <a name="display-data"></a>Anzeigen von Daten

In diesen Beispielen wird gezeigt, wie Sie mit Arbeitsblattdaten arbeiten und Benutzern eine bessere Ansicht oder Organisation bieten.

### <a name="apply-conditional-formatting"></a>Anwenden bedingter Formatierung

In diesem Beispiel wird die bedingte Formatierung auf den aktuell verwendeten Bereich im Arbeitsblatt angewendet. Die bedingte Formatierung ist eine grüne Füllung für die oberen 10% der Werte.

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

In diesem Beispiel wird die erste PivotTable in der Arbeitsmappe gesucht und die Werte in den Zellen "Gesamtsumme" (wie in der Abbildung unten grün hervorgehoben) protokolliert.

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

## <a name="scenario-samples"></a>Szenario-Beispiele

Beispiele für größere Lösungen aus der realen Welt finden Sie unter [Beispielszenarien für Office-Skripts](scenarios/sample-scenario-overview.md).

## <a name="suggest-new-samples"></a>Neue Beispiele vorschlagen

Wir begrüßen Vorschläge für neue Beispiele. Wenn es ein gängiges Szenario gibt, das anderen Skript Entwicklern helfen würde, teilen Sie uns dies im Abschnitt Feedback unten.
