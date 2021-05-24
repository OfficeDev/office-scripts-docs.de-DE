---
title: Verwenden von integrierten JavaScript-Objekten in Office-Skripts
description: Aufrufen von integrierten JavaScript-APIs aus einem Office skript in Excel im Web.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 680dd326e357bd06e2fc66cba5bd6745bbd33c24
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545047"
---
# <a name="use-built-in-javascript-objects-in-office-scripts"></a>Verwenden von integrierten JavaScript-Objekten in Office Skripts

JavaScript stellt mehrere integrierte Objekte zur Verfügung, die Sie in Ihren Office-Skripts verwenden können, unabhängig davon, ob Sie In JavaScript oder [TypeScript](../overview/code-editor-environment.md) (eine Übersatz von JavaScript) skripten. In diesem Artikel wird beschrieben, wie Sie einige der integrierten JavaScript-Objekte in Office Skripts für Excel im Web.

> [!NOTE]
> Eine vollständige Liste aller integrierten JavaScript-Objekte finden Sie im Artikel standard [built-in objects von](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) Mozilla.

## <a name="array"></a>Array

Das [Array-Objekt](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) bietet eine standardisierte Möglichkeit, mit Arrays in Ihrem Skript zu arbeiten. Arrays sind zwar Standard-JavaScript-Konstrukte, beziehen sich jedoch auf zwei Office Skripts: Bereiche und Auflistungen.

### <a name="work-with-ranges"></a>Arbeiten mit Bereichen

Bereiche enthalten mehrere zweidimensionale Arrays, die den Zellen in diesem Bereich direkt zuordnungen. Diese Arrays enthalten spezifische Informationen zu jeder Zelle in diesem Bereich. Gibt beispielsweise alle Werte in diesen Zellen zurück (mit den Zeilen und Spalten der zweidimensionalen Arrayzuordnung zu den Zeilen und Spalten dieses `Range.getValues` Arbeitsblattunterabschnitts). `Range.getFormulas` und `Range.getNumberFormats` sind andere häufig verwendete Methoden, die Arrays wie `Range.getValues` zurückgeben.

Das folgende Skript durchsucht den **A1:D4-Bereich** nach einem beliebigen Zahlenformat, das "$" enthält. Das Skript legt die Füllfarbe in diesen Zellen auf "gelb" fest.

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

### <a name="work-with-collections"></a>Arbeiten mit Sammlungen

Viele Excel sind in einer Auflistung enthalten. Die Auflistung wird von der Office Skripts-API verwaltet und als Array verfügbar gemacht. Beispielsweise sind alle [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in einem Arbeitsblatt in einer enthalten, die von `Shape[]` der Methode zurückgegeben `Worksheet.getShapes` wird. Sie können dieses Array verwenden, um Werte aus der Auflistung zu lesen, oder Sie können über die Methoden des übergeordneten Objekts auf bestimmte Objekte `get*` zugreifen.

> [!NOTE]
> Fügen Sie keine Objekte manuell aus diesen Auflistungsarrays hinzu oder entfernen Sie sie nicht. Verwenden Sie `add` die Methoden für die übergeordneten Objekte und die Methoden für die Objekte des `delete` Auflistungstyps. Fügen Sie z. B. ein [Table-Arbeitsblatt](/javascript/api/office-scripts/excelscript/excelscript.table) mit [der](/javascript/api/office-scripts/excelscript/excelscript.worksheet) `Worksheet.addTable` -Methode hinzu, und entfernen Sie `Table` die using `Table.delete` -Methode.

Das folgende Skript protokolliert den Typ jeder Form im aktuellen Arbeitsblatt.

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

Mit dem folgenden Skript wird die älteste Form im aktuellen Arbeitsblatt gelöscht.

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

Das [Date-Objekt](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) bietet eine standardisierte Möglichkeit, mit Datumsangaben in Ihrem Skript zu arbeiten. `Date.now()` generiert ein Objekt mit dem aktuellen Datum und der aktuellen Uhrzeit, was beim Hinzufügen von Zeitstempeln zum Dateneintrag Ihres Skripts hilfreich ist.

Das folgende Skript fügt dem Arbeitsblatt das aktuelle Datum hinzu. Beachten Sie, dass Excel methode den Wert als Datum erkennt und das Zahlenformat `toLocaleDateString` der Zelle automatisch ändert.

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

Der [Abschnitt Mit Datumsangaben](../resources/samples/excel-samples.md#dates) arbeiten in den Beispielen enthält mehr datumsbezogene Skripts.

## <a name="math"></a>Mathematik

Das [Math-Objekt](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) stellt Methoden und Konstanten für allgemeine mathematische Vorgänge zur Verfügung. Diese bieten viele Funktionen, die auch in Excel verfügbar sind, ohne das Berechnungsmodul der Arbeitsmappe verwenden zu müssen. Dadurch wird das Skript vor der Abfrage der Arbeitsmappe bewahrt, wodurch die Leistung verbessert wird.

Das folgende Skript verwendet, um die kleinste Zahl im `Math.min` **A1:D4-Bereich** zu finden und zu protokollieren. Beachten Sie, dass in diesem Beispiel davon ausgegangen wird, dass der gesamte Bereich nur Zahlen und keine Zeichenfolgen enthält.

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

Office Skripts unterstützen die Verwendung externer Drittanbieterbibliotheken nicht. Ihr Skript kann nur die integrierten JavaScript-Objekte und die Office Skript-APIs verwenden.

## <a name="see-also"></a>Siehe auch

- [Integrierte Standardobjekte](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Office Skripts-Code-Editor-Umgebung](../overview/code-editor-environment.md)
