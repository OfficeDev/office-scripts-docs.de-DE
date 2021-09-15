---
title: Verwenden von integrierten JavaScript-Objekten in Office-Skripts
description: Aufrufen integrierter JavaScript-APIs aus einem Office-Skript in Excel im Web.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: bf12a405814bb626a72c1de4f4c75462ce0018ec
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/15/2021
ms.locfileid: "59327689"
---
# <a name="use-built-in-javascript-objects-in-office-scripts"></a>Verwenden integrierter JavaScript-Objekte in Office Skripts

JavaScript bietet mehrere integrierte Objekte, die Sie in Ihren Office Skripts verwenden können, unabhängig davon, ob Sie Skripts in JavaScript oder [TypeScript](../overview/code-editor-environment.md) (eine Obermenge von JavaScript) erstellen. In diesem Artikel wird beschrieben, wie Sie einige der integrierten JavaScript-Objekte in Office Skripts für Excel im Web verwenden können.

> [!NOTE]
> Eine vollständige Liste aller integrierten JavaScript-Objekte finden Sie im Artikel zu den [integrierten Standardobjekten](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) von Mozilla.

## <a name="array"></a>Array

Das [Array-Objekt](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) bietet eine standardisierte Methode zum Arbeiten mit Arrays in Ihrem Skript. Arrays sind zwar Standard-JavaScript-Konstrukte, beziehen sich jedoch auf zwei Hauptmöglichkeiten auf Office Skripts: Bereiche und Sammlungen.

### <a name="work-with-ranges"></a>Arbeiten mit Bereichen

Bereiche enthalten mehrere zweidimensionale Arrays, die den Zellen in diesem Bereich direkt zugeordnet werden. Diese Arrays enthalten spezifische Informationen zu jeder Zelle in diesem Bereich. Gibt beispielsweise `Range.getValues` alle Werte in diesen Zellen zurück (mit den Zeilen und Spalten der zweidimensionalen Arrayzuordnung zu den Zeilen und Spalten dieses Arbeitsblatt-Unterabschnitts). `Range.getFormulas` und `Range.getNumberFormats` andere häufig verwendete Methoden sind, die Arrays wie `Range.getValues` .

Das folgende Skript durchsucht den **A1:D4-Bereich** nach einem beliebigen Zahlenformat, das ein "$" enthält. Das Skript legt die Füllfarbe in diesen Zellen auf "Gelb" fest.

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

Viele Excel-Objekte sind in einer Auflistung enthalten. Die Auflistung wird von der Office Scripts-API verwaltet und als Array verfügbar gemacht. Beispielsweise sind alle [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in einem Arbeitsblatt in einer `Shape[]` enthalten, die von der Methode zurückgegeben `Worksheet.getShapes` wird. Sie können dieses Array verwenden, um Werte aus der Auflistung zu lesen, oder Sie können über die Methoden des übergeordneten Objekts auf bestimmte Objekte `get*` zugreifen.

> [!NOTE]
> Fügen Sie diesen Sammlungsarrays keine Objekte manuell hinzu oder entfernen Sie diese. Verwenden Sie die `add` Methoden für die übergeordneten Objekte und die `delete` Methoden für die Auflistungstypobjekte. Fügen Sie z. [B.](/javascript/api/office-scripts/excelscript/excelscript.table) einem [Arbeitsblatt](/javascript/api/office-scripts/excelscript/excelscript.worksheet) mit der Methode eine Tabelle `Worksheet.addTable` hinzu, und entfernen Sie die `Table` using `Table.delete` -Methode.

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

Das [Date-Objekt](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) bietet eine standardisierte Methode zum Arbeiten mit Datumsangaben in Ihrem Skript. `Date.now()` generiert ein Objekt mit dem aktuellen Datum und der aktuellen Uhrzeit, was beim Hinzufügen von Zeitstempeln zur Dateneingabe Ihres Skripts hilfreich ist.

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

Der Abschnitt ["Arbeiten mit Datumsangaben"](../resources/samples/excel-samples.md#dates) in den Beispielen enthält weitere datumsbezogene Skripts.

## <a name="math"></a>Mathematik

Das [Math-Objekt](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) stellt Methoden und Konstanten für allgemeine mathematische Vorgänge bereit. Diese bieten viele Funktionen, die auch in Excel verfügbar sind, ohne dass das Berechnungsmodul der Arbeitsmappe verwendet werden muss. Dadurch wird verhindert, dass das Skript die Arbeitsmappe abfragen muss, was die Leistung verbessert.

Das folgende Skript `Math.min` wird verwendet, um die kleinste Zahl im **A1:D4-Bereich** zu suchen und zu protokollieren. Beachten Sie, dass in diesem Beispiel davon ausgegangen wird, dass der gesamte Bereich nur Zahlen und keine Zeichenfolgen enthält.

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

Office Skripts unterstützen nicht die Verwendung externer Drittanbieterbibliotheken. Ihr Skript kann nur die integrierten JavaScript-Objekte und die Office Scripts-APIs verwenden.

## <a name="see-also"></a>Siehe auch

- [Integrierte Standardobjekte](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Office Skript-Code-Editor-Umgebung](../overview/code-editor-environment.md)
