---
title: Verwenden von integrierten JavaScript-Objekten in Office-Skripts
description: So rufen Sie integrierte JavaScript-APIs aus einem Office Skript in Excel im Web auf.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 680dd326e357bd06e2fc66cba5bd6745bbd33c24
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545047"
---
# <a name="use-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="5c242-103">Verwenden integrierter JavaScript-Objekte in Office Skripts</span><span class="sxs-lookup"><span data-stu-id="5c242-103">Use built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="5c242-104">JavaScript stellt mehrere integrierte Objekte bereit, die Sie in Ihren Office Scripts verwenden können, unabhängig davon, ob Sie Skripts in JavaScript oder [TypeScript](../overview/code-editor-environment.md) (einer Übergruppe von JavaScript) erstellen.</span><span class="sxs-lookup"><span data-stu-id="5c242-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="5c242-105">In diesem Artikel wird beschrieben, wie Sie einige der integrierten JavaScript-Objekte in Office Scripts für Excel im Web verwenden können.</span><span class="sxs-lookup"><span data-stu-id="5c242-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="5c242-106">Eine vollständige Liste aller integrierten JavaScript-Objekte finden Sie im Artikel ["Standard- "Standard"](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) von Mozilla.</span><span class="sxs-lookup"><span data-stu-id="5c242-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="5c242-107">Array</span><span class="sxs-lookup"><span data-stu-id="5c242-107">Array</span></span>

<span data-ttu-id="5c242-108">Das [Array-Objekt](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) bietet eine standardisierte Möglichkeit, mit Arrays in Ihrem Skript zu arbeiten.</span><span class="sxs-lookup"><span data-stu-id="5c242-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="5c242-109">Obwohl Arrays Standard-JavaScript-Konstrukte sind, beziehen sie sich auf zwei Hauptarten auf Office Skripts: Bereiche und Auflistungen.</span><span class="sxs-lookup"><span data-stu-id="5c242-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="work-with-ranges"></a><span data-ttu-id="5c242-110">Arbeiten mit Reichweiten</span><span class="sxs-lookup"><span data-stu-id="5c242-110">Work with ranges</span></span>

<span data-ttu-id="5c242-111">Bereiche enthalten mehrere zweidimensionale Arrays, die direkt den Zellen in diesem Bereich zugeordnet werden.</span><span class="sxs-lookup"><span data-stu-id="5c242-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="5c242-112">Diese Arrays enthalten spezifische Informationen zu jeder Zelle in diesem Bereich.</span><span class="sxs-lookup"><span data-stu-id="5c242-112">These arrays contain specific information about each cell in that range.</span></span> <span data-ttu-id="5c242-113">Gibt z. B. `Range.getValues` alle Werte in diesen Zellen zurück (mit den Zeilen und Spalten des zweidimensionalen Arrays, die den Zeilen und Spalten dieses Arbeitsblattunterabschnitts zugeordnet sind).</span><span class="sxs-lookup"><span data-stu-id="5c242-113">For example, `Range.getValues` returns all the values in those cells (with the rows and columns of the two-dimensional array mapping to the rows and columns of that worksheet subsection).</span></span> <span data-ttu-id="5c242-114">`Range.getFormulas` und `Range.getNumberFormats` sind andere häufig verwendete Methoden, die Arrays wie `Range.getValues` zurückgeben.</span><span class="sxs-lookup"><span data-stu-id="5c242-114">`Range.getFormulas` and `Range.getNumberFormats` are other frequently used methods that return arrays like `Range.getValues`.</span></span>

<span data-ttu-id="5c242-115">Das folgende Skript durchsucht den **A1:D4-Bereich** nach einem beliebigen Zahlenformat, das ein "$" enthält.</span><span class="sxs-lookup"><span data-stu-id="5c242-115">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="5c242-116">Das Skript legt die Füllfarbe in diesen Zellen auf "gelb" fest.</span><span class="sxs-lookup"><span data-stu-id="5c242-116">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="work-with-collections"></a><span data-ttu-id="5c242-117">Arbeiten mit Sammlungen</span><span class="sxs-lookup"><span data-stu-id="5c242-117">Work with collections</span></span>

<span data-ttu-id="5c242-118">Viele Excel Objekte sind in einer Auflistung enthalten.</span><span class="sxs-lookup"><span data-stu-id="5c242-118">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="5c242-119">Die Auflistung wird von der Office Scripts-API verwaltet und als Array verfügbar gemacht.</span><span class="sxs-lookup"><span data-stu-id="5c242-119">The collection is managed by the Office Scripts API and exposed as an array.</span></span> <span data-ttu-id="5c242-120">Beispielsweise sind alle [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in einem Arbeitsblatt in einem `Shape[]` enthalten, das von der Methode zurückgegeben `Worksheet.getShapes` wird.</span><span class="sxs-lookup"><span data-stu-id="5c242-120">For example, all [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in a worksheet are contained in a `Shape[]` that is returned by the `Worksheet.getShapes` method.</span></span> <span data-ttu-id="5c242-121">Sie können dieses Array verwenden, um Werte aus der Auflistung zu lesen, oder Sie können auf bestimmte Objekte aus den Methoden des übergeordneten Objekts `get*` zugreifen.</span><span class="sxs-lookup"><span data-stu-id="5c242-121">You can use this array to read values from the collection, or you can access specific objects from the parent object's `get*` methods.</span></span>

> [!NOTE]
> <span data-ttu-id="5c242-122">Fügen Sie keine Objekte manuell zu diesen Auflistungsarrays hinzu oder entfernen Sie sie.</span><span class="sxs-lookup"><span data-stu-id="5c242-122">Do not manually add or remove objects from these collection arrays.</span></span> <span data-ttu-id="5c242-123">Verwenden Sie die `add` Methoden für die übergeordneten Objekte und die `delete` Methoden für die Auflistungsobjekte.</span><span class="sxs-lookup"><span data-stu-id="5c242-123">Use the `add` methods on the parent objects and the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="5c242-124">Fügen Sie beispielsweise eine [Tabelle](/javascript/api/office-scripts/excelscript/excelscript.table) zu einem [Arbeitsblatt](/javascript/api/office-scripts/excelscript/excelscript.worksheet) mit der `Worksheet.addTable` Methode hinzu, und entfernen Sie die `Table` verwendung `Table.delete` .</span><span class="sxs-lookup"><span data-stu-id="5c242-124">For example, add a [Table](/javascript/api/office-scripts/excelscript/excelscript.table) to a [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) with the `Worksheet.addTable` method and remove the `Table` using `Table.delete`.</span></span>

<span data-ttu-id="5c242-125">Das folgende Skript protokolliert den Typ jedes Shapes im aktuellen Arbeitsblatt.</span><span class="sxs-lookup"><span data-stu-id="5c242-125">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="5c242-126">Das folgende Skript löscht die älteste Form im aktuellen Arbeitsblatt.</span><span class="sxs-lookup"><span data-stu-id="5c242-126">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="5c242-127">Datum</span><span class="sxs-lookup"><span data-stu-id="5c242-127">Date</span></span>

<span data-ttu-id="5c242-128">Das [Date-Objekt](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) bietet eine standardisierte Möglichkeit, mit Datumsangaben in Ihrem Skript zu arbeiten.</span><span class="sxs-lookup"><span data-stu-id="5c242-128">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="5c242-129">`Date.now()` generiert ein Objekt mit dem aktuellen Datum und der aktuellen Uhrzeit, was nützlich ist, wenn Sie dem Dateneintrag Ihres Skripts Zeitstempel hinzufügen.</span><span class="sxs-lookup"><span data-stu-id="5c242-129">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="5c242-130">Das folgende Skript fügt dem Arbeitsblatt das aktuelle Datum hinzu.</span><span class="sxs-lookup"><span data-stu-id="5c242-130">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="5c242-131">Beachten Sie, dass Excel mit der `toLocaleDateString` Methode den Wert als Datum erkennt und das Zahlenformat der Zelle automatisch ändert.</span><span class="sxs-lookup"><span data-stu-id="5c242-131">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="5c242-132">Der Abschnitt [Arbeiten mit Datumsangaben](../resources/samples/excel-samples.md#dates) in den Beispielen enthält weitere datumsbezogene Skripts.</span><span class="sxs-lookup"><span data-stu-id="5c242-132">The [Work with dates](../resources/samples/excel-samples.md#dates) section of the samples has more date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="5c242-133">Mathematik</span><span class="sxs-lookup"><span data-stu-id="5c242-133">Math</span></span>

<span data-ttu-id="5c242-134">Das [Math-Objekt](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) stellt Methoden und Konstanten für allgemeine mathematische Operationen bereit.</span><span class="sxs-lookup"><span data-stu-id="5c242-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="5c242-135">Diese bieten viele Funktionen, die auch in Excel verfügbar sind, ohne dass das Berechnungsmodul der Arbeitsmappe verwendet werden muss.</span><span class="sxs-lookup"><span data-stu-id="5c242-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="5c242-136">Dadurch wird verhindert, dass Das Skript die Arbeitsmappe abfragen muss, was die Leistung verbessert.</span><span class="sxs-lookup"><span data-stu-id="5c242-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="5c242-137">Das folgende Skript `Math.min` verwendet, um die kleinste Zahl im **A1:D4-Bereich** zu suchen und zu protokollieren.</span><span class="sxs-lookup"><span data-stu-id="5c242-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="5c242-138">Beachten Sie, dass in diesem Beispiel davon ausgegangen wird, dass der gesamte Bereich nur Zahlen und keine Zeichenfolgen enthält.</span><span class="sxs-lookup"><span data-stu-id="5c242-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="5c242-139">Die Verwendung externer JavaScript-Bibliotheken wird nicht unterstützt</span><span class="sxs-lookup"><span data-stu-id="5c242-139">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="5c242-140">Office Skripts unterstützen nicht die Verwendung externer Bibliotheken von Drittanbietern.</span><span class="sxs-lookup"><span data-stu-id="5c242-140">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="5c242-141">Ihr Skript kann nur die integrierten JavaScript-Objekte und die Office Scripts-APIs verwenden.</span><span class="sxs-lookup"><span data-stu-id="5c242-141">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="5c242-142">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="5c242-142">See also</span></span>

- [<span data-ttu-id="5c242-143">Standard-Einbauobjekte</span><span class="sxs-lookup"><span data-stu-id="5c242-143">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="5c242-144">Office Skriptcode-Editor-Umgebung</span><span class="sxs-lookup"><span data-stu-id="5c242-144">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
