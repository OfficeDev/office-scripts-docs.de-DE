---
title: Verwenden von integrierten JavaScript-Objekten in Office-Skripts
description: Aufrufen von integrierten JavaScript-APIs aus einem Office-Skript in Excel im Internet.
ms.date: 04/06/2020
localization_priority: Normal
ms.openlocfilehash: a4b698215edea5f266e159fee0e08690904dd379
ms.sourcegitcommit: 0b2232c4c228b14d501edb8bb489fe0e84748b42
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/08/2020
ms.locfileid: "43191013"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a>Verwenden von integrierten JavaScript-Objekten in Office-Skripts

JavaScript stellt verschiedene integrierte Objekte bereit, die Sie in Ihren Office-Skripts verwenden können, unabhängig davon, ob Sie JavaScript-Skripts [oder Skript](../overview/code-editor-environment.md) (eine Obermenge von JavaScript). In diesem Artikel wird beschrieben, wie Sie einige der integrierten JavaScript-Objekte in Office-Skripts für Excel im Internet verwenden können.

> [!NOTE]
> Eine vollständige Liste aller integrierten JavaScript-Objekte finden Sie unter Mozillas [Standard mäßigen integrierten Objects-](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) Artikel.

## <a name="array"></a>Array

Das [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) -Objekt bietet eine standardisierte Möglichkeit zum Arbeiten mit Arrays in Ihrem Skript. Obwohl Arrays standardmäßige JavaScript-Konstrukte sind, beziehen Sie sich auf zwei Hauptmethoden auf Office-Skripts: Bereiche und Auflistungen.

### <a name="working-with-ranges"></a>Arbeiten mit Bereichen

Bereiche enthalten mehrere zweidimensionale Arrays, die den Zellen in diesem Bereich direkt zugeordnet werden. Dazu gehören Eigenschaften wie `values`, `formulas`und. `numberFormat` Array-Type-Eigenschaften müssen wie alle anderen Eigenschaften [geladen](scripting-fundamentals.md#sync-and-load) werden.

Das folgende Skript durchsucht den Bereich **a1: D4** für ein beliebiges Zahlenformat, das einen Wert vom Datentyp "$" enthält. Das Skript legt die Füllfarbe in diesen Zellen auf "gelb" fest.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range From A1 to D4.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");

  // Load the numberFormat property on the range.
  range.load("numberFormat");
  await context.sync();

  // Iterate through the arrays of rows and columns corresponding to those in the range.
  range.numberFormat.forEach((rowItem, rowIndex) => {
    range.numberFormat[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).format.fill.color = "yellow";
      }
    });
  });
}
```

### <a name="working-with-collections"></a>Arbeiten mit Auflistungen

Viele Excel-Objekte sind in einer Auflistung enthalten. Beispielsweise sind alle [Formen](/javascript/api/office-scripts/excel/excel.shape) in einem Arbeitsblatt in einer [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (als `Worksheet.shapes` Eigenschaft) enthalten. Jedes `*Collection` Objekt enthält eine `items` Eigenschaft, bei der es sich um ein Array handelt, in dem die Objekte in dieser Auflistung gespeichert werden. Dies kann wie ein normales JavaScript-Array behandelt werden, aber die Elemente in der Auflistung müssen zuerst geladen werden. Wenn Sie mit einer Eigenschaft für jedes Objekt in der Auflistung arbeiten müssen, verwenden Sie eine hierarchische ladeanweisung (`items/propertyName`).

Im folgenden Skript wird der Typ jedes Shapes im aktuellen Arbeitsblatt protokolliert.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.shapes;
  shapes.load("items/type");
  await context.sync();

  // Log the type of every shape in the collection.
  shapes.items.forEach((shape) => {
    console.log(shape.type);
  });
}
```

Sie können einzelne Objekte aus einer Auflistung mithilfe der `getItem` Methoden oder `getItemAt` laden. `getItem`Ruft ein Objekt mithilfe eines eindeutigen Bezeichners wie einen Namen ab (solche Namen werden häufig durch Ihr Skript angegeben). `getItemAt`Ruft ein Objekt mithilfe seines Index in der Auflistung ab. Auf einen der beiden Aufrufe muss ein `await context.sync();` Befehl folgen, bevor das Objekt verwendet werden kann.

Das folgende Skript löscht die älteste Form im aktuellen Arbeitsblatt.

```Typescript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.shapes.getItemAt(0);

  // Sync to load `shape` from the collection.
  await context.sync();

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a>Datum

Das [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) -Objekt bietet eine standardisierte Möglichkeit zum Arbeiten mit Datumsangaben in Ihrem Skript. `Date.now()`generiert ein Objekt mit dem aktuellen Datum und der aktuellen Uhrzeit, was hilfreich ist, wenn dem Dateneintrag des Skripts Timestamps hinzugefügt werden.

Mit dem folgenden Skript wird dem Arbeitsblatt das aktuelle Datum hinzugefügt. Beachten Sie, dass Excel `toLocaleDateString` mithilfe der Methode den Wert als Datum erkennt und das Zahlenformat der Zelle automatisch ändert.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range for cell A1.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.values = [[date.toLocaleDateString()]];
}
```

Der Abschnitt [Arbeiten mit Datumsangaben](../resources/excel-samples.md#work-with-dates) der Beispiele enthält mehr datumsbezogene Skripts.

## <a name="math"></a>Mathematik

Das [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) -Objekt stellt Methoden und Konstanten für allgemeine mathematische Vorgänge bereit. Diese stellen viele Funktionen bereit, die auch in Excel verfügbar sind, ohne dass das Berechnungsmodul der Arbeitsmappe verwendet werden muss. Dadurch wird das Skript so gespeichert, dass die Arbeitsmappe nicht abgefragt werden muss, wodurch die Leistung verbessert wird.

Das folgende Skript verwendet `Math.min` , um die kleinste Zahl im Bereich **a1: D4** zu suchen und zu protokollieren. Beachten Sie, dass in diesem Beispiel davon ausgegangen wird, dass der gesamte Bereich nur Zahlen, keine Zeichenfolgen enthält.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range from A1 to D4.
  let comparisonRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");
  
  // Load the range's values.
  comparisonRange.load("values");
  await context.sync();

  // Set the minimum values as the first value.
  let minimum = comparisonRange.values[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRange.values.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRange.values[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });
  
  console.log(minimum);
}

```

## <a name="see-also"></a>Siehe auch

- [Standard mäßige integrierte Objekte](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Code-Editor-Umgebung für Office-Skripts](../overview/code-editor-environment.md)
