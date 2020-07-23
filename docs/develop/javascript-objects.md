---
title: Verwenden von integrierten JavaScript-Objekten in Office-Skripts
description: Aufrufen von integrierten JavaScript-APIs aus einem Office-Skript in Excel im Internet.
ms.date: 07/16/2020
localization_priority: Normal
ms.openlocfilehash: 4bb5fb5444887005ececbbfdf0130cba3784e0c4
ms.sourcegitcommit: 8d549884e68170f808d3d417104a4451a37da83c
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/22/2020
ms.locfileid: "45229596"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a>Verwenden von integrierten JavaScript-Objekten in Office-Skripts

JavaScript stellt verschiedene integrierte Objekte bereit, die Sie in Ihren Office-Skripts verwenden können, unabhängig davon, ob Sie JavaScript-Skripts [oder Skript](../overview/code-editor-environment.md) (eine Obermenge von JavaScript). In diesem Artikel wird beschrieben, wie Sie einige der integrierten JavaScript-Objekte in Office-Skripts für Excel im Internet verwenden können.

> [!NOTE]
> Eine vollständige Liste aller integrierten JavaScript-Objekte finden Sie unter Mozillas [Standard mäßigen integrierten Objects-](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) Artikel.

## <a name="array"></a>Array

Das [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) -Objekt bietet eine standardisierte Möglichkeit zum Arbeiten mit Arrays in Ihrem Skript. Obwohl Arrays standardmäßige JavaScript-Konstrukte sind, beziehen Sie sich auf zwei Hauptmethoden auf Office-Skripts: Bereiche und Auflistungen.

### <a name="working-with-ranges"></a>Arbeiten mit Bereichen

Bereiche enthalten mehrere zweidimensionale Arrays, die den Zellen in diesem Bereich direkt zugeordnet werden. Diese Arrays enthalten spezifische Informationen zu jeder Zelle in diesem Bereich. Gibt beispielsweise `Range.getValues` alle Werte in diesen Zellen (mit den Zeilen und Spalten der zweidimensionalen Array Zuordnung zu den Zeilen und Spalten des Unterabschnitts dieses Arbeitsblatts) zurück. `Range.getFormulas`und `Range.getNumberFormats` sind andere häufig verwendete Methoden, die Arrays wie zurückgeben `Range.getValues` .

Das folgende Skript durchsucht den Bereich **a1: D4** für ein beliebiges Zahlenformat, das einen Wert vom Datentyp "$" enthält. Das Skript legt die Füllfarbe in diesen Zellen auf "gelb" fest.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range From A1 to D4.
  let range = workbook.getActiveWorksheet().getRange("A1:D4");

  // Get the number formats for each cell in the range.
  let rangeNumberFormats = range.getNumberFormats();
  // Iterate through the arrays of rows and columns corresponding to those in the range.
  rangeNumberFormats.forEach((rowItem, rowIndex) => {
    rangeNumberFormats[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).getFormat().getFill().setColor("yellow");
      }
    });
  });
}
```

### <a name="working-with-collections"></a>Arbeiten mit Auflistungen

Viele Excel-Objekte sind in einer Auflistung enthalten. Die Auflistung wird von der Office Scripts-API verwaltet und als Array verfügbar gemacht. Beispielsweise sind alle [Formen](/javascript/api/office-scripts/excelscript/excelscript.shape) in einem Arbeitsblatt in einer enthalten `Shape[]` , die von der-Methode zurückgegeben wird `Worksheet.getShapes` . Sie können dieses Array verwenden, um Werte aus der Auflistung zu lesen, oder Sie können von den Methoden des übergeordneten Objekts auf bestimmte Objekte zugreifen `get*` .

> [!NOTE]
> Fügen Sie keine Objekte dieser Auflistungs Arrays manuell hinzu oder entfernen Sie Sie. Verwenden Sie die `add` Methoden für die übergeordneten Objekte und die `delete` Methoden für die Auflistung-Type-Objekte. Fügen Sie beispielsweise eine [Tabelle](/javascript/api/office-scripts/excelscript/excelscript.table) mit der-Methode zu einem [Arbeitsblatt](/javascript/api/office-scripts/excelscript/excelscript.worksheet) hinzu, `Worksheet.addTable` und entfernen Sie die `Table` using `Table.delete` .

Im folgenden Skript wird der Typ jedes Shapes im aktuellen Arbeitsblatt protokolliert.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.getShapes();

  // Log the type of every shape in the collection.
  shapes.forEach((shape) => {
    console.log(shape.getType());
  });
}
```

Das folgende Skript löscht die älteste Form im aktuellen Arbeitsblatt.

```Typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.getShapes()[0];

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a>Datum

Das [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) -Objekt bietet eine standardisierte Möglichkeit zum Arbeiten mit Datumsangaben in Ihrem Skript. `Date.now()`generiert ein Objekt mit dem aktuellen Datum und der aktuellen Uhrzeit, was hilfreich ist, wenn dem Dateneintrag des Skripts Timestamps hinzugefügt werden.

Mit dem folgenden Skript wird dem Arbeitsblatt das aktuelle Datum hinzugefügt. Beachten Sie, dass Excel mithilfe der `toLocaleDateString` Methode den Wert als Datum erkennt und das Zahlenformat der Zelle automatisch ändert.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range for cell A1.
  let range = workbook.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.setValue(date.toLocaleDateString());
}
```

Der Abschnitt [Arbeiten mit Datumsangaben](../resources/excel-samples.md#dates) der Beispiele enthält mehr datumsbezogene Skripts.

## <a name="math"></a>Mathematik

Das [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) -Objekt stellt Methoden und Konstanten für allgemeine mathematische Vorgänge bereit. Diese stellen viele Funktionen bereit, die auch in Excel verfügbar sind, ohne dass das Berechnungsmodul der Arbeitsmappe verwendet werden muss. Dadurch wird das Skript so gespeichert, dass die Arbeitsmappe nicht abgefragt werden muss, wodurch die Leistung verbessert wird.

Das folgende Skript verwendet `Math.min` , um die kleinste Zahl im Bereich **a1: D4** zu suchen und zu protokollieren. Beachten Sie, dass in diesem Beispiel davon ausgegangen wird, dass der gesamte Bereich nur Zahlen, keine Zeichenfolgen enthält.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range from A1 to D4.
  let comparisonRange = workbook.getActiveWorksheet().getRange("A1:D4");

  // Load the range's values.
  let comparisonRangeValues = comparisonRange.getValues();

  // Set the minimum values as the first value.
  let minimum = comparisonRangeValues[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRangeValues.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRangeValues[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });

  console.log(minimum);
}

```

## <a name="use-of-external-javascript-libraries-is-not-supported"></a>Die Verwendung externer JavaScript-Bibliotheken wird nicht unterstützt.

Office-Skripts unterstützen nicht die Verwendung externer Bibliotheken von Drittanbietern. Ihr Skript kann nur die integrierten JavaScript-Objekte und die Office Scripts-APIs verwenden.

## <a name="see-also"></a>Siehe auch

- [Standard mäßige integrierte Objekte](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Code-Editor-Umgebung für Office-Skripts](../overview/code-editor-environment.md)
