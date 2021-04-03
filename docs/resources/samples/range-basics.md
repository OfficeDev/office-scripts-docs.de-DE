---
title: Grundlegendes zu Bereichsfunktionen in Office-Skripts
description: Erfahren Sie mehr über die Verwendung des Range-Objekts in Office-Skripts.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: 73eeba086aace6262c624de9074ffb301f6532bd
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571304"
---
# <a name="range-basics"></a><span data-ttu-id="4ed4c-103">Grundlagen des Bereichs</span><span class="sxs-lookup"><span data-stu-id="4ed4c-103">Range basics</span></span>

<span data-ttu-id="4ed4c-104">`Range` ist das Basisobjekt innerhalb des Office Scripts Excel-Objektmodells.</span><span class="sxs-lookup"><span data-stu-id="4ed4c-104">`Range` is the foundational object within the Office Scripts Excel object model.</span></span> <span data-ttu-id="4ed4c-105">[Bereichs-APIs](/javascript/api/office-scripts/excelscript/excelscript.range) ermöglichen den Zugriff auf daten und das Format, die im Raster verfügbar sind, und verknüpfen andere wichtige Objekte in Excel, z. B. Arbeitsblätter, Tabellen, Diagramme usw.</span><span class="sxs-lookup"><span data-stu-id="4ed4c-105">[Range APIs](/javascript/api/office-scripts/excelscript/excelscript.range) allow access to both data and format available on the grid and link other key objects within Excel such as worksheets, tables, charts, etc.</span></span>

<span data-ttu-id="4ed4c-106">Ein Bereich wird anhand seiner Adresse wie "A1:B4" oder mithilfe eines benannten Elements identifiziert, bei dem es sich um einen benannten Schlüssel für eine bestimmte Gruppe von Zellen handelt.</span><span class="sxs-lookup"><span data-stu-id="4ed4c-106">A range is identified using its address such as "A1:B4" or using a named-item, which is a named key for a given set of cells.</span></span> <span data-ttu-id="4ed4c-107">Im Excel-Objektmodell werden sowohl eine Zelle als auch eine Gruppe von Zellen als Bereich _bezeichnet._</span><span class="sxs-lookup"><span data-stu-id="4ed4c-107">In the Excel object model, both a cell and group of cells are referred as _range_.</span></span> <span data-ttu-id="4ed4c-108">`Range` kann Attribute auf Zellenebene wie Daten innerhalb einer Zelle sowie Attribute auf Zellen- und Zellenebene wie Format, Rahmen usw. enthalten.</span><span class="sxs-lookup"><span data-stu-id="4ed4c-108">`Range` can contain cell-level attributes such as data within a cell and also cell and cells-level attributes such as format, borders, etc.</span></span>

<span data-ttu-id="4ed4c-109">`Range` kann auch über die Auswahl des Benutzers, die aus mindestens einer Zelle besteht, erhalten werden.</span><span class="sxs-lookup"><span data-stu-id="4ed4c-109">`Range` can also be obtained via user's selection that consists of at least one cell.</span></span> <span data-ttu-id="4ed4c-110">Bei der Interaktion mit dem Bereich ist es wichtig, diese Zell- und Bereichsbeziehungen klar zu halten.</span><span class="sxs-lookup"><span data-stu-id="4ed4c-110">As you interact with the range, it's important to keep these cell and range relationships clear.</span></span>

<span data-ttu-id="4ed4c-111">Im Folgenden finden Sie die wichtigsten Getter, Setter und andere nützliche Methoden, die am häufigsten in Skripts verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="4ed4c-111">Following are the core set of getters, setters, and other useful methods most often used in scripts.</span></span> <span data-ttu-id="4ed4c-112">Dies ist ein hervorragender Ausgangspunkt für Ihre API-Reise.</span><span class="sxs-lookup"><span data-stu-id="4ed4c-112">This is a great starting point for your API journey.</span></span> <span data-ttu-id="4ed4c-113">Die späteren Abschnitte gruppieren die Methoden und helfen beim Erstellen eines mentalen Modells, während Sie beginnen, die APIs des `Range` Objekts zu entsperren.</span><span class="sxs-lookup"><span data-stu-id="4ed4c-113">The later sections group the methods and help to build a mental model as you begin to unlock the `Range` object's APIs.</span></span>

## <a name="example-scripts"></a><span data-ttu-id="4ed4c-114">Beispielskripts</span><span class="sxs-lookup"><span data-stu-id="4ed4c-114">Example scripts</span></span>

* [<span data-ttu-id="4ed4c-115">Grundlegendes Lesen und Schreiben</span><span class="sxs-lookup"><span data-stu-id="4ed4c-115">Basic read and write</span></span>](#basic-read-and-write)
* [<span data-ttu-id="4ed4c-116">Zeile am Ende des Arbeitsblatts hinzufügen</span><span class="sxs-lookup"><span data-stu-id="4ed4c-116">Add row at the end of worksheet</span></span>](#add-row-at-the-end-of-worksheet)
* [<span data-ttu-id="4ed4c-117">Spaltenfilter löschen</span><span class="sxs-lookup"><span data-stu-id="4ed4c-117">Clear column filter</span></span>](clear-table-filter-for-active-cell.md)
* [<span data-ttu-id="4ed4c-118">Jede Zelle mit einer eindeutigen Farbe färben</span><span class="sxs-lookup"><span data-stu-id="4ed4c-118">Color each cell with unique color</span></span>](#color-each-cell-with-unique-color)
* [<span data-ttu-id="4ed4c-119">Aktualisieren des Bereichs mit Werten mithilfe eines 2D-Arrays</span><span class="sxs-lookup"><span data-stu-id="4ed4c-119">Update range with values using 2-dimensional (2D) array</span></span>](#update-range-with-values-using-2d-array)

### <a name="basic-read-and-write"></a><span data-ttu-id="4ed4c-120">Grundlegendes Lesen und Schreiben</span><span class="sxs-lookup"><span data-stu-id="4ed4c-120">Basic read and write</span></span>

```TypeScript
/**
 * This script demonstrates basic read-write operations on the Range object.
 */
function main(workbook: ExcelScript.Workbook) {
  const cell = workbook.getActiveCell();
  const prevValue = cell.getValue();
  if (prevValue) {
      console.log(`Active cell's value is: ${prevValue}`);
  } else {
      console.log("Setting active cell's value..");
      cell.setValue("Sample");
  }

  // Get cell next to the right column and set its value and fill color.
  const nextCell = cell.getOffsetRange(0,1);
  nextCell.setValue("Next cell");
  console.log(`Next cell's address is: ${nextCell.getAddress()}`);
  console.log("Setting fill color and font color of next cell...");
  nextCell.getFormat().getFill().setColor("Magenta");
  nextCell.getFormat().getFill().setColor("Cyan");

  // Get the target range address to update with 2-dimensional value.
  const dataRange = nextCell.getOffsetRange(1, 0).getResizedRange(2, 1);
  const DATA = [
    [10, 7],
    [8, 15],
    [12, 1]
  ];
  console.log(`Updating range ${dataRange.getAddress()} with values: ${DATA}`);
  dataRange.setValues(DATA);

  // Formula range.
  const formulaRange = dataRange.getOffsetRange(3, 0).getRow(0);
  console.log(`Updating formula for range: ${formulaRange.getAddress()}`)
  // Since relative formula is being set, we can set the formula of the entire range to the same value.
  formulaRange.setFormulaR1C1("=SUM(R[-3]C:R[-1]C)");
  console.log(`Updating number format for range: ${formulaRange.getAddress()}`)
  // Since the number format is common to the entire range, we can set it to a common format.
  formulaRange.setNumberFormat("0.00");
  return;
}
```

### <a name="add-row-at-the-end-of-worksheet"></a><span data-ttu-id="4ed4c-121">Zeile am Ende des Arbeitsblatts hinzufügen</span><span class="sxs-lookup"><span data-stu-id="4ed4c-121">Add row at the end of worksheet</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet('Sheet5');
    const data = ['2016', 'Bikes', 'Seats', '1500', .05];
    addRow(sheet, data);
    return;
}

function addRow(sheet: ExcelScript.Worksheet, data: (string | number | boolean)[]): void {

    const usedRange = sheet.getUsedRange();
    let startCell: ExcelScript.Range;
    // If the sheet is empty, then use A1 as starting cell for the update.
    if (usedRange) {
      startCell = usedRange.getLastRow().getCell(0, 0).getOffsetRange(1, 0);
    } else {
      startCell = sheet.getRange('A1');
    }
    console.log(startCell.getAddress());
    const targetRange = startCell.getResizedRange(0, data.length - 1);
    targetRange.setValues([data]);
    return;
}
```

### <a name="color-each-cell-with-unique-color"></a><span data-ttu-id="4ed4c-122">Jede Zelle mit einer eindeutigen Farbe färben</span><span class="sxs-lookup"><span data-stu-id="4ed4c-122">Color each cell with unique color</span></span>

```TypeScript
/**
 * This sample demonstrates how to iterate over a selected range and set cell property.
   It colors each cell within the selected range with a random color.
 */
function main(workbook: ExcelScript.Workbook) {

    const syncStart = new Date().getTime();
    // Get selected range
    const range = workbook.getSelectedRange();
    const rows = range.getRowCount();
    const cols = range.getColumnCount();
    console.log("Start");

    // Color each cell with random color.
    for (let row = 0; row < rows; row++) {
        for (let col = 0; col < cols; col++) {
            range
                .getCell(row, col)
                .getFormat()
                .getFill()
                .setColor(`#${Math.random().toString(16).substr(-6)}`);
        }
    }

    console.log("End");
    const syncEnd = new Date().getTime();
    console.log("Completed, took: " + (syncEnd - syncStart) / 1000 + " Sec");
}
```

### <a name="update-range-with-values-using-2d-array"></a><span data-ttu-id="4ed4c-123">Aktualisieren des Bereichs mit Werten mithilfe des 2D-Arrays</span><span class="sxs-lookup"><span data-stu-id="4ed4c-123">Update range with values using 2D array</span></span>

<span data-ttu-id="4ed4c-124">Berechnet dynamisch die zu aktualisierende Bereichsdimension basierend auf 2D-Arraywerten.</span><span class="sxs-lookup"><span data-stu-id="4ed4c-124">Dynamically calculates the range dimension to update based on 2D array values.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const currentCell = workbook.getActiveCell();
  let inputRange = computeTargetRange(currentCell, DATA);
  // Set range values.
  console.log(inputRange.getAddress());
  inputRange.setValues(DATA);
  // Call a helper function to place border around the range.
  borderAround(inputRange);
}

/**
 * A helper function that computes the target range given the target range's starting cell and selected range. 
 */
function computeTargetRange(targetCell: ExcelScript.Range, data: string[][]): ExcelScript.Range {
  const targetRange = targetCell.getResizedRange(data.length - 1, data[0].length - 1);
  return targetRange;
}

/**
 * A helper function that places a border around the range.
 */
function borderAround(range: ExcelScript.Range): void {
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.dash);
  return;
}

// Values used for range setup.
const DATA = [
  ['Item', 'Bread', 'Donuts', 'Cookies', 'Cakes', 'Pies'],
  ['Amount', '2', '1.5', '4', '12', '26']
]
```

## <a name="training-videos-range-basics"></a><span data-ttu-id="4ed4c-125">Schulungsvideos: Grundlegendes zum Bereich</span><span class="sxs-lookup"><span data-stu-id="4ed4c-125">Training videos: Range basics</span></span>

<span data-ttu-id="4ed4c-126">_Grundlagen des Bereichs_</span><span class="sxs-lookup"><span data-stu-id="4ed4c-126">_Range basics_</span></span>

<span data-ttu-id="4ed4c-127">[![Schritt-für-Schritt-Video zu Range Basics ansehen](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Schritt-für-Schritt-Video zu Range-Grundlagen")</span><span class="sxs-lookup"><span data-stu-id="4ed4c-127">[![Watch step-by-step video on Range basics](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Step-by-step video on Range basics")</span></span>

<span data-ttu-id="4ed4c-128">_Zeile am Ende des Arbeitsblatts hinzufügen_</span><span class="sxs-lookup"><span data-stu-id="4ed4c-128">_Add row at the end of worksheet_</span></span>

<span data-ttu-id="4ed4c-129">[![Schritt-für-Schritt-Video zum Hinzufügen einer Zeile am Ende eines Arbeitsblatts ansehen](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "Schrittweises Video zum Hinzufügen einer Zeile am Ende eines Arbeitsblatts")</span><span class="sxs-lookup"><span data-stu-id="4ed4c-129">[![Watch step-by-step video on how to add a row at the end of a worksheet](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "Step-by-step video on how to add a row at the end of a worksheet")</span></span>

## <a name="methods-that-return-some-range-metadata"></a><span data-ttu-id="4ed4c-130">Methoden, die einige Bereichsmetadaten zurückgeben</span><span class="sxs-lookup"><span data-stu-id="4ed4c-130">Methods that return some range metadata</span></span>

* <span data-ttu-id="4ed4c-131">getAddress(), getAddressLocal()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-131">getAddress(), getAddressLocal()</span></span>
* <span data-ttu-id="4ed4c-132">getCellCount()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-132">getCellCount()</span></span>
* <span data-ttu-id="4ed4c-133">getRowCount(), getColumnCount()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-133">getRowCount(), getColumnCount()</span></span>

## <a name="methods-that-return-dataconstants-associated-with-a-given-range"></a><span data-ttu-id="4ed4c-134">Methoden, die Daten/Konstanten zurückgeben, die einem bestimmten Bereich zugeordnet sind</span><span class="sxs-lookup"><span data-stu-id="4ed4c-134">Methods that return data/constants associated with a given range</span></span>

### <a name="returned-as-single-cell-value"></a><span data-ttu-id="4ed4c-135">Als Einzelzellenwert zurückgegeben</span><span class="sxs-lookup"><span data-stu-id="4ed4c-135">Returned as single cell value</span></span>

* <span data-ttu-id="4ed4c-136">getFormula(), getFormulaLocal()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-136">getFormula(), getFormulaLocal()</span></span>
* <span data-ttu-id="4ed4c-137">getFormulaR1C1()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-137">getFormulaR1C1()</span></span>
* <span data-ttu-id="4ed4c-138">getNumberFormat(), getNumberFormatLocal()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-138">getNumberFormat(), getNumberFormatLocal()</span></span>
* <span data-ttu-id="4ed4c-139">getText()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-139">getText()</span></span>
* <span data-ttu-id="4ed4c-140">getValue()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-140">getValue()</span></span>
* <span data-ttu-id="4ed4c-141">getValueType()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-141">getValueType()</span></span>

### <a name="returned-as-2d-arrays-whole-range"></a><span data-ttu-id="4ed4c-142">Als 2D-Arrays zurückgegeben (ganzer Bereich)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-142">Returned as 2D arrays (whole range)</span></span>

* <span data-ttu-id="4ed4c-143">getFormulas(), getFormulasLocal()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-143">getFormulas(), getFormulasLocal()</span></span>
* <span data-ttu-id="4ed4c-144">getFormulasR1C1()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-144">getFormulasR1C1()</span></span>
* <span data-ttu-id="4ed4c-145">getNumberFormatCategories()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-145">getNumberFormatCategories()</span></span>
* <span data-ttu-id="4ed4c-146">getNumberFormats(), getNumberFormatsLocal()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-146">getNumberFormats(), getNumberFormatsLocal()</span></span>
* <span data-ttu-id="4ed4c-147">getTexts()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-147">getTexts()</span></span>
* <span data-ttu-id="4ed4c-148">getValues()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-148">getValues()</span></span>
* <span data-ttu-id="4ed4c-149">getValueTypes()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-149">getValueTypes()</span></span>
* <span data-ttu-id="4ed4c-150">getHidden()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-150">getHidden()</span></span>
* <span data-ttu-id="4ed4c-151">getIsEntireRow()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-151">getIsEntireRow()</span></span>
* <span data-ttu-id="4ed4c-152">getIsEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-152">getIsEntireColumn()</span></span>

## <a name="methods-that-return-other-range-object"></a><span data-ttu-id="4ed4c-153">Methoden, die ein anderes Bereichsobjekt zurückgeben</span><span class="sxs-lookup"><span data-stu-id="4ed4c-153">Methods that return other range object</span></span>

* <span data-ttu-id="4ed4c-154">getSurroundingRegion() –- ähnlich wie CurrentRegion in VBA</span><span class="sxs-lookup"><span data-stu-id="4ed4c-154">getSurroundingRegion() -- similar to CurrentRegion in VBA</span></span>
* <span data-ttu-id="4ed4c-155">getCell(row, column)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-155">getCell(row, column)</span></span>
* <span data-ttu-id="4ed4c-156">getColumn(column)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-156">getColumn(column)</span></span>
* <span data-ttu-id="4ed4c-157">getColumnHidden()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-157">getColumnHidden()</span></span>
* <span data-ttu-id="4ed4c-158">getColumnsAfter(count)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-158">getColumnsAfter(count)</span></span>
* <span data-ttu-id="4ed4c-159">getColumnsBefore(count)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-159">getColumnsBefore(count)</span></span>
* <span data-ttu-id="4ed4c-160">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-160">getEntireColumn()</span></span>
* <span data-ttu-id="4ed4c-161">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-161">getEntireRow()</span></span>
* <span data-ttu-id="4ed4c-162">getLastCell()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-162">getLastCell()</span></span>
* <span data-ttu-id="4ed4c-163">getLastColumn()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-163">getLastColumn()</span></span>
* <span data-ttu-id="4ed4c-164">getLastRow()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-164">getLastRow()</span></span>
* <span data-ttu-id="4ed4c-165">getRow(row)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-165">getRow(row)</span></span>
* <span data-ttu-id="4ed4c-166">getRowHidden()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-166">getRowHidden()</span></span>
* <span data-ttu-id="4ed4c-167">getRowsAbove(count)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-167">getRowsAbove(count)</span></span>
* <span data-ttu-id="4ed4c-168">getRowsBelow(count)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-168">getRowsBelow(count)</span></span>

<span data-ttu-id="4ed4c-169">**Wichtig/Interessant**</span><span class="sxs-lookup"><span data-stu-id="4ed4c-169">**Important/Interesting**</span></span>

* <span data-ttu-id="4ed4c-170">_Arbeitsmappe_.getSelectedRange()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-170">_workbook_.getSelectedRange()</span></span>
* <span data-ttu-id="4ed4c-171">_Arbeitsmappe_.getActiveCell()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-171">_workbook_.getActiveCell()</span></span>
* <span data-ttu-id="4ed4c-172">getUsedRange(valuesOnly)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-172">getUsedRange(valuesOnly)</span></span>
* <span data-ttu-id="4ed4c-173">getAbsoluteResizedRange(numRows, numColumns)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-173">getAbsoluteResizedRange(numRows, numColumns)</span></span>
* <span data-ttu-id="4ed4c-174">getOffsetRange(rowOffset, columnOffset)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-174">getOffsetRange(rowOffset, columnOffset)</span></span>
* <span data-ttu-id="4ed4c-175">getResizedRange(deltaRows, deltaColumns)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-175">getResizedRange(deltaRows, deltaColumns)</span></span>

## <a name="methods-that-return-a-range-object-in-relation-to-another-range-object"></a><span data-ttu-id="4ed4c-176">Methoden, die ein Bereichsobjekt in Bezug auf ein anderes Bereichsobjekt zurückgeben</span><span class="sxs-lookup"><span data-stu-id="4ed4c-176">Methods that return a range object in relation to another range object</span></span>

* <span data-ttu-id="4ed4c-177">getBoundingRect(anotherRange)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-177">getBoundingRect(anotherRange)</span></span>
* <span data-ttu-id="4ed4c-178">getIntersection(anotherRange)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-178">getIntersection(anotherRange)</span></span>

## <a name="methods-that-return-other-objects-non-range-objects"></a><span data-ttu-id="4ed4c-179">Methoden, die andere Objekte zurückgeben (Nichtbereichsobjekte)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-179">Methods that return other objects (non-range objects)</span></span>

* <span data-ttu-id="4ed4c-180">getDirectPrecedents()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-180">getDirectPrecedents()</span></span>
* <span data-ttu-id="4ed4c-181">getWorksheet()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-181">getWorksheet()</span></span>
* <span data-ttu-id="4ed4c-182">getTables(fullyContained)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-182">getTables(fullyContained)</span></span>
* <span data-ttu-id="4ed4c-183">getPivotTables(fullyContained)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-183">getPivotTables(fullyContained)</span></span>
* <span data-ttu-id="4ed4c-184">getDataValidation()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-184">getDataValidation()</span></span>
* <span data-ttu-id="4ed4c-185">getPredefinedCellStyle()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-185">getPredefinedCellStyle()</span></span>

## <a name="set-methods"></a><span data-ttu-id="4ed4c-186">Festlegen von Methoden</span><span class="sxs-lookup"><span data-stu-id="4ed4c-186">Set methods</span></span>

### <a name="singular-cell-set-methods"></a><span data-ttu-id="4ed4c-187">Singular Cell Set-Methoden</span><span class="sxs-lookup"><span data-stu-id="4ed4c-187">Singular cell set methods</span></span>

* <span data-ttu-id="4ed4c-188">setFormula(formula)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-188">setFormula(formula)</span></span>
* <span data-ttu-id="4ed4c-189">setFormulaLocal(formulaLocal)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-189">setFormulaLocal(formulaLocal)</span></span>
* <span data-ttu-id="4ed4c-190">setFormulaR1C1(formulaR1C1)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-190">setFormulaR1C1(formulaR1C1)</span></span>
* <span data-ttu-id="4ed4c-191">setNumberFormatLocal(numberFormatLocal)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-191">setNumberFormatLocal(numberFormatLocal)</span></span>
* <span data-ttu-id="4ed4c-192">setValue(value)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-192">setValue(value)</span></span>

### <a name="2d--entire-range-set-methods"></a><span data-ttu-id="4ed4c-193">2D/gesamte Bereichssatzmethoden</span><span class="sxs-lookup"><span data-stu-id="4ed4c-193">2D / entire range set methods</span></span>

* <span data-ttu-id="4ed4c-194">setFormulas(formulas)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-194">setFormulas(formulas)</span></span>
* <span data-ttu-id="4ed4c-195">setFormulasLocal(formulasLocal)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-195">setFormulasLocal(formulasLocal)</span></span>
* <span data-ttu-id="4ed4c-196">setFormulasR1C1(formulasR1C1)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-196">setFormulasR1C1(formulasR1C1)</span></span>
* <span data-ttu-id="4ed4c-197">setNumberFormat(numberFormat)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-197">setNumberFormat(numberFormat)</span></span>
* <span data-ttu-id="4ed4c-198">setNumberFormats(numberFormats)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-198">setNumberFormats(numberFormats)</span></span>
* <span data-ttu-id="4ed4c-199">setNumberFormatsLocal(numberFormatsLocal)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-199">setNumberFormatsLocal(numberFormatsLocal)</span></span>
* <span data-ttu-id="4ed4c-200">setValues(values)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-200">setValues(values)</span></span>

## <a name="other-methods"></a><span data-ttu-id="4ed4c-201">Andere Methoden</span><span class="sxs-lookup"><span data-stu-id="4ed4c-201">Other methods</span></span>

* <span data-ttu-id="4ed4c-202">merge(across)</span><span class="sxs-lookup"><span data-stu-id="4ed4c-202">merge(across)</span></span>
* <span data-ttu-id="4ed4c-203">unmerge()</span><span class="sxs-lookup"><span data-stu-id="4ed4c-203">unmerge()</span></span>

## <a name="coming-soon"></a><span data-ttu-id="4ed4c-204">In Kürze verfügbar</span><span class="sxs-lookup"><span data-stu-id="4ed4c-204">Coming soon</span></span>

* <span data-ttu-id="4ed4c-205">Bereichs-Edge-APIs</span><span class="sxs-lookup"><span data-stu-id="4ed4c-205">Range edge APIs</span></span>
