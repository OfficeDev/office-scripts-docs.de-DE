---
title: Verschieben von Zeilen über Tabellen mithilfe von Office-Skripts
description: Erfahren Sie, wie Sie Zeilen über Tabellen hinweg verschieben, indem Sie Filter speichern, dann die Filter verarbeiten und erneut anwenden.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: a7c28c4fef91402b8889d749a03f3aab5e615521
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572619"
---
# <a name="move-rows-across-tables"></a>Verschieben von Zeilen über Tabellen hinweg

In diesem Skript werden folgende Schritte ausgeführt:

* Wählt Zeilen aus der Quelltabelle aus, in denen der Wert in einer Spalte einem Wert (`FILTER_VALUE` im Skript) entspricht.
* Verschiebt alle markierten Zeilen in die Zieltabelle in einem anderen Arbeitsblatt.
* Wendet die relevanten Filter erneut auf die Quelltabelle an.

## <a name="sample-excel-file"></a>Excel-Beispieldatei

Laden Sie die Datei [input-table-filters.xlsx](input-table-filters.xlsx) für eine sofort einsatzbereite Arbeitsmappe herunter. Fügen Sie das folgende Skript hinzu, um das Beispiel selbst auszuprobieren!

## <a name="sample-code-move-rows-using-range-values"></a>Beispielcode: Verschieben von Zeilen mithilfe von Bereichswerten

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // You can change these names to match the data in your workbook.
  const TARGET_TABLE_NAME = 'Table1';
  const SOURCE_TABLE_NAME = 'Table2';

  // Select what will be moved between tables.
  const FILTER_COLUMN_INDEX = 1;
  const FILTER_VALUE = 'Clothing';

  // Get the Table objects.
  let targetTable = workbook.getTable(TARGET_TABLE_NAME);
  let sourceTable = workbook.getTable(SOURCE_TABLE_NAME);

  // If either table is missing, report that information and stop the script.
  if (!targetTable || !sourceTable) {
    console.log(`Tables missing - Check to make sure both source (${TARGET_TABLE_NAME}) and target table (${SOURCE_TABLE_NAME}) are present before running the script. `);
    return;
  }

  // Save the filter criteria currently on the source table.
  const originalTableFilters = {};
  // For each table column, collect the filter criteria on that column.
  sourceTable.getColumns().forEach((column) => {
    let originalColumnFilter = column.getFilter().getCriteria();
    if (originalColumnFilter) {
      originalTableFilters[column.getName()] = originalColumnFilter;
    }
  });

  // Get all the data from the table.
  const sourceRange = sourceTable.getRangeBetweenHeaderAndTotal();
  const dataRows: (number | string | boolean)[][] = sourceTable.getRangeBetweenHeaderAndTotal().getValues();

  // Create variables to hold the rows to be moved and their addresses.
  let rowsToMoveValues: (number | string | boolean)[][] = [];
  let rowAddressToRemove: string[] = [];

  // Get the data values from the source table.
  for (let i = 0; i < dataRows.length; i++) { 
    if (dataRows[i][FILTER_COLUMN_INDEX] === FILTER_VALUE) {
      rowsToMoveValues.push(dataRows[i]);

      // Get the intersection between table address and the entire row where we found the match. This provides the address of the range to remove.
      let address = sourceRange.getIntersection(sourceRange.getCell(i,0).getEntireRow()).getAddress();
      rowAddressToRemove.push(address);
    }
  }

  // If there are no data rows to process, end the script.
  if (rowsToMoveValues.length < 1) {
    console.log('No rows selected from the source table match the filter criteria.');
    return;
  }

  console.log(`Adding ${rowsToMoveValues.length} rows to target table.`);

  // Insert rows at the end of target table.
  targetTable.addRows(-1, rowsToMoveValues)

  // Remove the rows from the source table.
  const sheet = sourceTable.getWorksheet();

  // Remove all filters before removing rows.
  sourceTable.getAutoFilter().clearCriteria();

  // Important: Remove the rows starting at the bottom of the table.
  // Otherwise, the lower rows change position before they are deleted.
  console.log(`Removing ${rowAddressToRemove.length} rows from the source table.`);
  rowAddressToRemove.reverse().forEach((address) => {
    sheet.getRange(address).delete(ExcelScript.DeleteShiftDirection.up);
  });

  // Reapply the original filters. 
  Object.keys(originalTableFilters).forEach((columnName) => {
      sourceTable.getColumnByName(columnName).getFilter().apply(originalTableFilters[columnName]);
    });
}
```

## <a name="training-video-move-rows-across-tables"></a>Schulungsvideo: Verschieben von Zeilen über Tabellen hinweg

[Sehen Sie sich Sudhi Ramamurthy bei diesem Beispiel auf YouTube an](https://youtu.be/_3t3Pk4i2L0). Es gibt zwei Skripts, die in der Videolösung gezeigt werden. Der Hauptunterschied besteht darin, wie die Zeilen ausgewählt werden.

* In der ersten Variante werden die Zeilen durch Anwenden des Tabellenfilters und Lesen des sichtbaren Bereichs ausgewählt.
* In der zweiten Zeile werden die Zeilen ausgewählt, indem die Werte gelesen und die Zeilenwerte extrahiert werden (was im Beispiel auf dieser Seite verwendet wird).
