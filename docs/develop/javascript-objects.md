---
title: Verwenden integrierter JavaScript-Objekte in Office-Skripts
description: Aufrufen von integrierten JavaScript-APIs aus einem Office-Skript in Excel im Internet.
ms.date: 01/21/2020
localization_priority: Normal
ms.openlocfilehash: e0fcd98117125ead18e55675e195415ff59c0c5d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700202"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="a8005-103">Verwenden integrierter JavaScript-Objekte in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="a8005-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="a8005-104">JavaScript stellt verschiedene integrierte Objekte bereit, die Sie in Ihren Office-Skripts verwenden können, unabhängig davon, ob Sie JavaScript-Skripts [oder Skript](../overview/code-editor-environment.md) (eine Obermenge von JavaScript).</span><span class="sxs-lookup"><span data-stu-id="a8005-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="a8005-105">In diesem Artikel wird beschrieben, wie Sie einige der integrierten JavaScript-Objekte in Office-Skripts für Excel im Internet verwenden können.</span><span class="sxs-lookup"><span data-stu-id="a8005-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="a8005-106">Eine vollständige Liste aller integrierten JavaScript-Objekte finden Sie unter Mozillas [Standard mäßigen integrierten Objects-](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) Artikel.</span><span class="sxs-lookup"><span data-stu-id="a8005-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="a8005-107">Array</span><span class="sxs-lookup"><span data-stu-id="a8005-107">Array</span></span>

<span data-ttu-id="a8005-108">Das [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) -Objekt bietet eine standardisierte Möglichkeit zum Arbeiten mit Arrays in Ihrem Skript.</span><span class="sxs-lookup"><span data-stu-id="a8005-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="a8005-109">Obwohl Arrays standardmäßige JavaScript-Konstrukte sind, beziehen Sie sich auf zwei Hauptmethoden auf Office-Skripts: Bereiche und Auflistungen.</span><span class="sxs-lookup"><span data-stu-id="a8005-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="a8005-110">Arbeiten mit Bereichen</span><span class="sxs-lookup"><span data-stu-id="a8005-110">Working with ranges</span></span>

<span data-ttu-id="a8005-111">Bereiche enthalten mehrere zweidimensionale Arrays, die den Zellen in diesem Bereich direkt zugeordnet werden.</span><span class="sxs-lookup"><span data-stu-id="a8005-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="a8005-112">Dazu gehören Eigenschaften wie `values`, `formulas`und. `numberFormat`</span><span class="sxs-lookup"><span data-stu-id="a8005-112">These include properties such as `values`, `formulas`, and `numberFormat`.</span></span> <span data-ttu-id="a8005-113">Array-Type-Eigenschaften müssen wie alle anderen Eigenschaften [geladen](scripting-fundamentals.md#sync-and-load) werden.</span><span class="sxs-lookup"><span data-stu-id="a8005-113">Array-type properties must be [loaded](scripting-fundamentals.md#sync-and-load) like any other properties.</span></span>

<span data-ttu-id="a8005-114">Das folgende Skript durchsucht den Bereich **a1: D4** für ein beliebiges Zahlenformat, das einen Wert vom Datentyp "$" enthält.</span><span class="sxs-lookup"><span data-stu-id="a8005-114">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="a8005-115">Das Skript legt die Füllfarbe in diesen Zellen auf "gelb" fest.</span><span class="sxs-lookup"><span data-stu-id="a8005-115">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="working-with-collections"></a><span data-ttu-id="a8005-116">Arbeiten mit Auflistungen</span><span class="sxs-lookup"><span data-stu-id="a8005-116">Working with collections</span></span>

<span data-ttu-id="a8005-117">Viele Excel-Objekte sind in einer Auflistung enthalten.</span><span class="sxs-lookup"><span data-stu-id="a8005-117">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="a8005-118">Beispielsweise sind alle [Formen](/javascript/api/office-scripts/excel/excel.shape) in einem Arbeitsblatt in einer [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (als `Worksheet.shapes` Eigenschaft) enthalten.</span><span class="sxs-lookup"><span data-stu-id="a8005-118">For example, all [Shapes](/javascript/api/office-scripts/excel/excel.shape) in a worksheet are contained in a [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (as the `Worksheet.shapes` property).</span></span> <span data-ttu-id="a8005-119">Jedes `*Collection` Objekt enthält eine `items` Eigenschaft, bei der es sich um ein Array handelt, in dem die Objekte in dieser Auflistung gespeichert werden.</span><span class="sxs-lookup"><span data-stu-id="a8005-119">Each `*Collection` object contains an `items` property, which is an array that stores the objects inside that collection.</span></span> <span data-ttu-id="a8005-120">Dies kann wie ein normales JavaScript-Array behandelt werden, aber die Elemente in der Auflistung müssen zuerst geladen werden.</span><span class="sxs-lookup"><span data-stu-id="a8005-120">This can be treated like a normal JavaScript array, but the items in the collection have to first be loaded.</span></span> <span data-ttu-id="a8005-121">Wenn Sie mit einer Eigenschaft für jedes Objekt in der Auflistung arbeiten müssen, verwenden Sie eine hierarchische ladeanweisung (`items/propertyName`).</span><span class="sxs-lookup"><span data-stu-id="a8005-121">If you need to work with a property on every object in the collection, use a hierarchal load statement (`items/propertyName`).</span></span>

<span data-ttu-id="a8005-122">Im folgenden Skript wird der Typ jedes Shapes im aktuellen Arbeitsblatt protokolliert.</span><span class="sxs-lookup"><span data-stu-id="a8005-122">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="a8005-123">Sie können einzelne Objekte aus einer Auflistung mithilfe der `getItem` Methoden oder `getItemAt` laden.</span><span class="sxs-lookup"><span data-stu-id="a8005-123">You can load individual objects from a collection using the `getItem` or `getItemAt` methods.</span></span> <span data-ttu-id="a8005-124">`getItem`Ruft ein Objekt mithilfe eines eindeutigen Bezeichners wie einen Namen ab (solche Namen werden häufig durch Ihr Skript angegeben).</span><span class="sxs-lookup"><span data-stu-id="a8005-124">`getItem` gets an object by using a unique identifier like a name (such names are often specified by your script).</span></span> <span data-ttu-id="a8005-125">`getItemAt`Ruft ein Objekt mithilfe seines Index in der Auflistung ab.</span><span class="sxs-lookup"><span data-stu-id="a8005-125">`getItemAt` gets an object by using its index in the collection.</span></span> <span data-ttu-id="a8005-126">Auf einen der beiden Aufrufe muss ein `await context.sync();` Befehl folgen, bevor das Objekt verwendet werden kann.</span><span class="sxs-lookup"><span data-stu-id="a8005-126">Either call must be followed by a `await context.sync();` command before the object can be used.</span></span>

<span data-ttu-id="a8005-127">Das folgende Skript löscht die älteste Form im aktuellen Arbeitsblatt.</span><span class="sxs-lookup"><span data-stu-id="a8005-127">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="a8005-128">Datum</span><span class="sxs-lookup"><span data-stu-id="a8005-128">Date</span></span>

<span data-ttu-id="a8005-129">Das [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) -Objekt bietet eine standardisierte Möglichkeit zum Arbeiten mit Datumsangaben in Ihrem Skript.</span><span class="sxs-lookup"><span data-stu-id="a8005-129">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="a8005-130">`Date.now()`generiert ein Objekt mit dem aktuellen Datum und der aktuellen Uhrzeit, was hilfreich ist, wenn dem Dateneintrag des Skripts Timestamps hinzugefügt werden.</span><span class="sxs-lookup"><span data-stu-id="a8005-130">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="a8005-131">Mit dem folgenden Skript wird dem Arbeitsblatt das aktuelle Datum hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="a8005-131">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="a8005-132">Beachten Sie, dass Excel `toLocaleDateString` mithilfe der Methode den Wert als Datum erkennt und das Zahlenformat der Zelle automatisch ändert.</span><span class="sxs-lookup"><span data-stu-id="a8005-132">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

## <a name="math"></a><span data-ttu-id="a8005-133">Mathematik</span><span class="sxs-lookup"><span data-stu-id="a8005-133">Math</span></span>

<span data-ttu-id="a8005-134">Das [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) -Objekt stellt Methoden und Konstanten für allgemeine mathematische Vorgänge bereit.</span><span class="sxs-lookup"><span data-stu-id="a8005-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="a8005-135">Diese stellen viele Funktionen bereit, die auch in Excel verfügbar sind, ohne dass das Berechnungsmodul der Arbeitsmappe verwendet werden muss.</span><span class="sxs-lookup"><span data-stu-id="a8005-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="a8005-136">Dadurch wird das Skript so gespeichert, dass die Arbeitsmappe nicht abgefragt werden muss, wodurch die Leistung verbessert wird.</span><span class="sxs-lookup"><span data-stu-id="a8005-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="a8005-137">Das folgende Skript verwendet `Math.min` , um die kleinste Zahl im Bereich **a1: D4** zu suchen und zu protokollieren.</span><span class="sxs-lookup"><span data-stu-id="a8005-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="a8005-138">Beachten Sie, dass in diesem Beispiel davon ausgegangen wird, dass der gesamte Bereich nur Zahlen, keine Zeichenfolgen enthält.</span><span class="sxs-lookup"><span data-stu-id="a8005-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="a8005-139">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="a8005-139">See also</span></span>

- [<span data-ttu-id="a8005-140">Standard mäßige integrierte Objekte</span><span class="sxs-lookup"><span data-stu-id="a8005-140">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="a8005-141">Code-Editor-Umgebung für Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="a8005-141">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
