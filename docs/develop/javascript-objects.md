---
title: Verwenden von integrierten JavaScript-Objekten in Office-Skripts
description: Aufrufen von integrierten JavaScript-APIs aus einem Office-Skript in Excel im Internet.
ms.date: 06/29/2020
localization_priority: Normal
ms.openlocfilehash: 1c8ac757574e8c4be64b373f8d4bf421ddfa0c79
ms.sourcegitcommit: 30750c4392db3ef057075a5702abb92863c93eda
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/01/2020
ms.locfileid: "44999260"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="5ac3b-103">Verwenden von integrierten JavaScript-Objekten in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="5ac3b-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="5ac3b-104">JavaScript stellt verschiedene integrierte Objekte bereit, die Sie in Ihren Office-Skripts verwenden können, unabhängig davon, ob Sie JavaScript-Skripts [oder Skript](../overview/code-editor-environment.md) (eine Obermenge von JavaScript).</span><span class="sxs-lookup"><span data-stu-id="5ac3b-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="5ac3b-105">In diesem Artikel wird beschrieben, wie Sie einige der integrierten JavaScript-Objekte in Office-Skripts für Excel im Internet verwenden können.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="5ac3b-106">Eine vollständige Liste aller integrierten JavaScript-Objekte finden Sie unter Mozillas [Standard mäßigen integrierten Objects-](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) Artikel.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="5ac3b-107">Array</span><span class="sxs-lookup"><span data-stu-id="5ac3b-107">Array</span></span>

<span data-ttu-id="5ac3b-108">Das [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) -Objekt bietet eine standardisierte Möglichkeit zum Arbeiten mit Arrays in Ihrem Skript.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="5ac3b-109">Obwohl Arrays standardmäßige JavaScript-Konstrukte sind, beziehen Sie sich auf zwei Hauptmethoden auf Office-Skripts: Bereiche und Auflistungen.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="5ac3b-110">Arbeiten mit Bereichen</span><span class="sxs-lookup"><span data-stu-id="5ac3b-110">Working with ranges</span></span>

<span data-ttu-id="5ac3b-111">Bereiche enthalten mehrere zweidimensionale Arrays, die den Zellen in diesem Bereich direkt zugeordnet werden.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="5ac3b-112">Diese Arrays enthalten spezifische Informationen zu jeder Zelle in diesem Bereich.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-112">These arrays contain specific information about each cell in that range.</span></span> <span data-ttu-id="5ac3b-113">Gibt beispielsweise `Range.getValues` alle Werte in diesen Zellen (mit den Zeilen und Spalten der zweidimensionalen Array Zuordnung zu den Zeilen und Spalten des Unterabschnitts dieses Arbeitsblatts) zurück.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-113">For example, `Range.getValues` returns all the values in those cells (with the rows and columns of the two-dimensional array mapping to the rows and columns of that worksheet subsection).</span></span> <span data-ttu-id="5ac3b-114">`Range.getFormulas`und `Range.getNumberFormats` sind andere häufig verwendete Methoden, die Arrays wie zurückgeben `Range.getValues` .</span><span class="sxs-lookup"><span data-stu-id="5ac3b-114">`Range.getFormulas` and `Range.getNumberFormats` are other frequently used methods that return arrays like `Range.getValues`.</span></span>

<span data-ttu-id="5ac3b-115">Das folgende Skript durchsucht den Bereich **a1: D4** für ein beliebiges Zahlenformat, das einen Wert vom Datentyp "$" enthält.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-115">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="5ac3b-116">Das Skript legt die Füllfarbe in diesen Zellen auf "gelb" fest.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-116">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="working-with-collections"></a><span data-ttu-id="5ac3b-117">Arbeiten mit Auflistungen</span><span class="sxs-lookup"><span data-stu-id="5ac3b-117">Working with collections</span></span>

<span data-ttu-id="5ac3b-118">Viele Excel-Objekte sind in einer Auflistung enthalten.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-118">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="5ac3b-119">Die Auflistung wird von der Office Scripts-API verwaltet und als Array verfügbar gemacht.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-119">The collection is managed by the Office Scripts API and exposed as an array.</span></span> <span data-ttu-id="5ac3b-120">Beispielsweise sind alle [Formen](/javascript/api/office-scripts/excelscript/excelscript.shape) in einem Arbeitsblatt in einer enthalten `Shape[]` , die von der-Methode zurückgegeben wird `Worksheet.getShapes` .</span><span class="sxs-lookup"><span data-stu-id="5ac3b-120">For example, all [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in a worksheet are contained in a `Shape[]` that is returned by the `Worksheet.getShapes` method.</span></span> <span data-ttu-id="5ac3b-121">Sie können dieses Array verwenden, um Werte aus der Auflistung zu lesen, oder Sie können von den Methoden des übergeordneten Objekts auf bestimmte Objekte zugreifen `get*` .</span><span class="sxs-lookup"><span data-stu-id="5ac3b-121">You can use this array to read values from the collection, or you can access specific objects from the parent object's `get*` methods.</span></span>

> [!NOTE]
> <span data-ttu-id="5ac3b-122">Fügen Sie keine Objekte dieser Auflistungs Arrays manuell hinzu oder entfernen Sie Sie.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-122">Do not manually add or remove objects from these collection arrays.</span></span> <span data-ttu-id="5ac3b-123">Verwenden Sie die `add` Methoden für die übergeordneten Objekte und die `delete` Methoden für die Auflistung-Type-Objekte.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-123">Use the `add` methods on the parent objects and the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="5ac3b-124">Fügen Sie beispielsweise eine [Tabelle](/javascript/api/office-scripts/excelscript/excelscript.table) mit der-Methode zu einem [Arbeitsblatt](/javascript/api/office-scripts/excelscript/excelscript.worksheet) hinzu, `Worksheet.addTable` und entfernen Sie die `Table` using `Table.delete` .</span><span class="sxs-lookup"><span data-stu-id="5ac3b-124">For example, add a [Table](/javascript/api/office-scripts/excelscript/excelscript.table) to a [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) with the `Worksheet.addTable` method and remove the `Table` using `Table.delete`.</span></span>

<span data-ttu-id="5ac3b-125">Im folgenden Skript wird der Typ jedes Shapes im aktuellen Arbeitsblatt protokolliert.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-125">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="5ac3b-126">Das folgende Skript löscht die älteste Form im aktuellen Arbeitsblatt.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-126">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="5ac3b-127">Datum</span><span class="sxs-lookup"><span data-stu-id="5ac3b-127">Date</span></span>

<span data-ttu-id="5ac3b-128">Das [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) -Objekt bietet eine standardisierte Möglichkeit zum Arbeiten mit Datumsangaben in Ihrem Skript.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-128">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="5ac3b-129">`Date.now()`generiert ein Objekt mit dem aktuellen Datum und der aktuellen Uhrzeit, was hilfreich ist, wenn dem Dateneintrag des Skripts Timestamps hinzugefügt werden.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-129">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="5ac3b-130">Mit dem folgenden Skript wird dem Arbeitsblatt das aktuelle Datum hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-130">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="5ac3b-131">Beachten Sie, dass Excel mithilfe der `toLocaleDateString` Methode den Wert als Datum erkennt und das Zahlenformat der Zelle automatisch ändert.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-131">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="5ac3b-132">Der Abschnitt [Arbeiten mit Datumsangaben](../resources/excel-samples.md#work-with-dates) der Beispiele enthält mehr datumsbezogene Skripts.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-132">The [Work with dates](../resources/excel-samples.md#work-with-dates) section of the samples has more Date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="5ac3b-133">Mathematik</span><span class="sxs-lookup"><span data-stu-id="5ac3b-133">Math</span></span>

<span data-ttu-id="5ac3b-134">Das [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) -Objekt stellt Methoden und Konstanten für allgemeine mathematische Vorgänge bereit.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="5ac3b-135">Diese stellen viele Funktionen bereit, die auch in Excel verfügbar sind, ohne dass das Berechnungsmodul der Arbeitsmappe verwendet werden muss.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="5ac3b-136">Dadurch wird das Skript so gespeichert, dass die Arbeitsmappe nicht abgefragt werden muss, wodurch die Leistung verbessert wird.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="5ac3b-137">Das folgende Skript verwendet `Math.min` , um die kleinste Zahl im Bereich **a1: D4** zu suchen und zu protokollieren.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="5ac3b-138">Beachten Sie, dass in diesem Beispiel davon ausgegangen wird, dass der gesamte Bereich nur Zahlen, keine Zeichenfolgen enthält.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="5ac3b-139">Die Verwendung externer JavaScript-Bibliotheken wird nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-139">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="5ac3b-140">Office-Skripts unterstützen nicht die Verwendung externer Bibliotheken von Drittanbietern.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-140">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="5ac3b-141">Ihr Skript kann nur die integrierten JavaScript-Objekte und die Office Scripts-APIs verwenden.</span><span class="sxs-lookup"><span data-stu-id="5ac3b-141">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="5ac3b-142">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="5ac3b-142">See also</span></span>

- [<span data-ttu-id="5ac3b-143">Standard mäßige integrierte Objekte</span><span class="sxs-lookup"><span data-stu-id="5ac3b-143">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="5ac3b-144">Code-Editor-Umgebung für Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="5ac3b-144">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
