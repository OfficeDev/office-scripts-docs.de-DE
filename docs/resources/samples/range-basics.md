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
# <a name="range-basics"></a>Grundlagen des Bereichs

`Range` ist das Basisobjekt innerhalb des Office Scripts Excel-Objektmodells. [Bereichs-APIs](/javascript/api/office-scripts/excelscript/excelscript.range) ermöglichen den Zugriff auf daten und das Format, die im Raster verfügbar sind, und verknüpfen andere wichtige Objekte in Excel, z. B. Arbeitsblätter, Tabellen, Diagramme usw.

Ein Bereich wird anhand seiner Adresse wie "A1:B4" oder mithilfe eines benannten Elements identifiziert, bei dem es sich um einen benannten Schlüssel für eine bestimmte Gruppe von Zellen handelt. Im Excel-Objektmodell werden sowohl eine Zelle als auch eine Gruppe von Zellen als Bereich _bezeichnet._ `Range` kann Attribute auf Zellenebene wie Daten innerhalb einer Zelle sowie Attribute auf Zellen- und Zellenebene wie Format, Rahmen usw. enthalten.

`Range` kann auch über die Auswahl des Benutzers, die aus mindestens einer Zelle besteht, erhalten werden. Bei der Interaktion mit dem Bereich ist es wichtig, diese Zell- und Bereichsbeziehungen klar zu halten.

Im Folgenden finden Sie die wichtigsten Getter, Setter und andere nützliche Methoden, die am häufigsten in Skripts verwendet werden. Dies ist ein hervorragender Ausgangspunkt für Ihre API-Reise. Die späteren Abschnitte gruppieren die Methoden und helfen beim Erstellen eines mentalen Modells, während Sie beginnen, die APIs des `Range` Objekts zu entsperren.

## <a name="example-scripts"></a>Beispielskripts

* [Grundlegendes Lesen und Schreiben](#basic-read-and-write)
* [Zeile am Ende des Arbeitsblatts hinzufügen](#add-row-at-the-end-of-worksheet)
* [Spaltenfilter löschen](clear-table-filter-for-active-cell.md)
* [Jede Zelle mit einer eindeutigen Farbe färben](#color-each-cell-with-unique-color)
* [Aktualisieren des Bereichs mit Werten mithilfe eines 2D-Arrays](#update-range-with-values-using-2d-array)

### <a name="basic-read-and-write"></a>Grundlegendes Lesen und Schreiben

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

### <a name="add-row-at-the-end-of-worksheet"></a>Zeile am Ende des Arbeitsblatts hinzufügen

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

### <a name="color-each-cell-with-unique-color"></a>Jede Zelle mit einer eindeutigen Farbe färben

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

### <a name="update-range-with-values-using-2d-array"></a>Aktualisieren des Bereichs mit Werten mithilfe des 2D-Arrays

Berechnet dynamisch die zu aktualisierende Bereichsdimension basierend auf 2D-Arraywerten.

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

## <a name="training-videos-range-basics"></a>Schulungsvideos: Grundlegendes zum Bereich

_Grundlagen des Bereichs_

[![Schritt-für-Schritt-Video zu Range Basics ansehen](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Schritt-für-Schritt-Video zu Range-Grundlagen")

_Zeile am Ende des Arbeitsblatts hinzufügen_

[![Schritt-für-Schritt-Video zum Hinzufügen einer Zeile am Ende eines Arbeitsblatts ansehen](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "Schrittweises Video zum Hinzufügen einer Zeile am Ende eines Arbeitsblatts")

## <a name="methods-that-return-some-range-metadata"></a>Methoden, die einige Bereichsmetadaten zurückgeben

* getAddress(), getAddressLocal()
* getCellCount()
* getRowCount(), getColumnCount()

## <a name="methods-that-return-dataconstants-associated-with-a-given-range"></a>Methoden, die Daten/Konstanten zurückgeben, die einem bestimmten Bereich zugeordnet sind

### <a name="returned-as-single-cell-value"></a>Als Einzelzellenwert zurückgegeben

* getFormula(), getFormulaLocal()
* getFormulaR1C1()
* getNumberFormat(), getNumberFormatLocal()
* getText()
* getValue()
* getValueType()

### <a name="returned-as-2d-arrays-whole-range"></a>Als 2D-Arrays zurückgegeben (ganzer Bereich)

* getFormulas(), getFormulasLocal()
* getFormulasR1C1()
* getNumberFormatCategories()
* getNumberFormats(), getNumberFormatsLocal()
* getTexts()
* getValues()
* getValueTypes()
* getHidden()
* getIsEntireRow()
* getIsEntireColumn()

## <a name="methods-that-return-other-range-object"></a>Methoden, die ein anderes Bereichsobjekt zurückgeben

* getSurroundingRegion() –- ähnlich wie CurrentRegion in VBA
* getCell(row, column)
* getColumn(column)
* getColumnHidden()
* getColumnsAfter(count)
* getColumnsBefore(count)
* getEntireColumn()
* getEntireRow()
* getLastCell()
* getLastColumn()
* getLastRow()
* getRow(row)
* getRowHidden()
* getRowsAbove(count)
* getRowsBelow(count)

**Wichtig/Interessant**

* _Arbeitsmappe_.getSelectedRange()
* _Arbeitsmappe_.getActiveCell()
* getUsedRange(valuesOnly)
* getAbsoluteResizedRange(numRows, numColumns)
* getOffsetRange(rowOffset, columnOffset)
* getResizedRange(deltaRows, deltaColumns)

## <a name="methods-that-return-a-range-object-in-relation-to-another-range-object"></a>Methoden, die ein Bereichsobjekt in Bezug auf ein anderes Bereichsobjekt zurückgeben

* getBoundingRect(anotherRange)
* getIntersection(anotherRange)

## <a name="methods-that-return-other-objects-non-range-objects"></a>Methoden, die andere Objekte zurückgeben (Nichtbereichsobjekte)

* getDirectPrecedents()
* getWorksheet()
* getTables(fullyContained)
* getPivotTables(fullyContained)
* getDataValidation()
* getPredefinedCellStyle()

## <a name="set-methods"></a>Festlegen von Methoden

### <a name="singular-cell-set-methods"></a>Singular Cell Set-Methoden

* setFormula(formula)
* setFormulaLocal(formulaLocal)
* setFormulaR1C1(formulaR1C1)
* setNumberFormatLocal(numberFormatLocal)
* setValue(value)

### <a name="2d--entire-range-set-methods"></a>2D/gesamte Bereichssatzmethoden

* setFormulas(formulas)
* setFormulasLocal(formulasLocal)
* setFormulasR1C1(formulasR1C1)
* setNumberFormat(numberFormat)
* setNumberFormats(numberFormats)
* setNumberFormatsLocal(numberFormatsLocal)
* setValues(values)

## <a name="other-methods"></a>Andere Methoden

* merge(across)
* unmerge()

## <a name="coming-soon"></a>In Kürze verfügbar

* Bereichs-Edge-APIs
