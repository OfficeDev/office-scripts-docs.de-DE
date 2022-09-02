---
title: Kombinieren von Daten aus mehreren Excel-Tabellen in einer einzelnen Tabelle
description: Erfahren Sie, wie Sie Mithilfe von Office-Skripts Daten aus mehreren Excel-Tabellen in einer einzigen Tabelle kombinieren.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3db510514c676b9012fd47abc2a7e92492a9cf87
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572451"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a>Kombinieren von Daten aus mehreren Excel-Tabellen in einer einzelnen Tabelle

In diesem Beispiel werden Daten aus mehreren Excel-Tabellen in einer einzigen Tabelle kombiniert, die alle Zeilen enthält. Es wird davon ausgegangen, dass alle verwendeten Tabellen dieselbe Struktur aufweisen.

Es gibt zwei Varianten dieses Skripts:

1. Das [erste Skript](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) kombiniert alle Tabellen in der Excel-Datei.
1. Das [zweite Skript](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) ruft selektiv Tabellen innerhalb einer Gruppe von Arbeitsblättern ab.

## <a name="sample-excel-file"></a>Excel-Beispieldatei

Laden Sie [tables-copy.xlsx](tables-copy.xlsx) für eine sofort einsatzbereite Arbeitsmappe herunter. Fügen Sie die folgenden Skripts hinzu, um das Beispiel selbst auszuprobieren!

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Beispielcode: Kombinieren von Daten aus mehreren Excel-Tabellen in einer einzigen Tabelle

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Delete the "Combined" worksheet, if it's present.
  workbook.getWorksheet('Combined')?.delete();

  // Create a new worksheet named "Combined" for the combined table.
  const newSheet = workbook.addWorksheet('Combined');
  
  // Get the header values for the first table in the workbook.
  // This also saves the table list before we add the new, combined table.
  const tables = workbook.getTables();    
  const headerValues = tables[0].getHeaderRowRange().getTexts();
  console.log(headerValues);

  // Copy the headers on a new worksheet to an equal-sized range.
  const targetRange = newSheet.getRange('A1').getResizedRange(headerValues.length-1, headerValues[0].length-1);
  targetRange.setValues(headerValues);

  // Add the data from each table in the workbook to the new table.
  const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
  for (let table of tables) {      
    let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
    let rowCount = table.getRowCount();

    // If the table is not empty, add its rows to the combined table.
    if (rowCount > 0) {
      combinedTable.addRows(-1, dataValues);
    }
  }
}
```

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a>Beispielcode: Kombinieren von Daten aus mehreren Excel-Tabellen in ausgewählten Arbeitsblättern in einer einzigen Tabelle

Laden Sie die Beispieldatei [tables-select-copy.xlsx](tables-select-copy.xlsx) herunter, und verwenden Sie sie mit dem folgenden Skript, um sie selbst auszuprobieren!

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Set the worksheet names to get tables from.
  const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    
  // Delete the "Combined" worksheet, if it's present.
  workbook.getWorksheet('Combined')?.delete();

  // Create a new worksheet named "Combined" for the combined table.
  const newSheet = workbook.addWorksheet('Combined');

  // Create a new table with the same headers as the other tables.
  const headerValues = workbook.getWorksheet(sheetNames[0]).getTables()[0].getHeaderRowRange().getTexts();
  const targetRange = newSheet.getRange('A1').getResizedRange(headerValues.length-1, headerValues[0].length-1);
  targetRange.setValues(headerValues);
  const combinedTable = newSheet.addTable(targetRange.getAddress(), true);

  // Go through each listed worksheet and get their tables.
  sheetNames.forEach((sheet) => {
    const tables = workbook.getWorksheet(sheet).getTables();     
    for (let table of tables) {
      // Get the rows from the tables.
      let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
      let rowCount = table.getRowCount();

      // If there's data in the table, add it to the combined table.
      if (rowCount > 0) {
          combinedTable.addRows(-1, dataValues);
      }
    }
  });
}
```

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Schulungsvideo: Kombinieren von Daten aus mehreren Excel-Tabellen in einer einzigen Tabelle

[Sehen Sie sich Sudhi Ramamurthy bei diesem Beispiel auf YouTube an](https://youtu.be/di-8JukK3Lc).
