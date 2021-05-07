---
title: Kombinieren von Daten aus Excel Tabellen in einer einzelnen Tabelle
description: Erfahren Sie, wie Sie Office Skripts verwenden, um Daten aus mehreren tabellen Excel in einer einzelnen Tabelle zu kombinieren.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: ac8c7d0a3f0f4f3d7d3217ffac31aff1a5595d17
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232445"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a>Kombinieren von Daten aus Excel Tabellen in einer einzelnen Tabelle

In diesem Beispiel werden Daten aus Excel Tabellen in einer einzigen Tabelle kombiniert, die alle Zeilen enthält. Es wird davon ausgegangen, dass alle verwendeten Tabellen dieselbe Struktur haben.

Es gibt zwei Variationen dieses Skripts:

1. Das [erste Skript](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) kombiniert alle Tabellen in der Excel Datei.
1. Das [zweite Skript](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) ruft tabellen selektiv innerhalb einer Gruppe von Arbeitsblättern ab.

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Beispielcode: Kombinieren von Daten aus Excel Tabellen in einer einzelnen Tabelle

Laden Sie die Beispieldatei <a href="tables-copy.xlsx">tables-copy.xlsx</a> und verwenden Sie sie mit dem folgenden Skript, um sie selbst auszuprobieren!

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    
    const tables = workbook.getTables();    
    const headerValues = tables[0].getHeaderRowRange().getTexts();
    console.log(headerValues);
    const targetRange = updateRange(newSheet, headerValues);
    const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
    for (let table of tables) {      
      let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
      let rowCount = table.getRowCount();
      if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
      }
    }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a>Beispielcode: Kombinieren von Daten aus mehreren Excel tabellen in ausgewählten Arbeitsblättern in einer einzelnen Tabelle

Laden Sie die Beispieldatei <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> und verwenden Sie sie mit dem folgenden Skript, um sie selbst auszuprobieren!

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    let targetTableCreated = false;
    let combinedTable;
    sheetNames.forEach((sheet) => {
      const tables = workbook.getWorksheet(sheet).getTables();
      if (!targetTableCreated) {
        const headerValues = tables[0].getHeaderRowRange().getTexts();
        const targetRange = updateRange(newSheet, headerValues);
        combinedTable = newSheet.addTable(targetRange.getAddress(), true);
        targetTableCreated = true;
      }      
      for (let table of tables) {
        let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
        let rowCount = table.getRowCount();
        if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
        }
      }
    })
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Schulungsvideo: Kombinieren von Daten aus mehreren Excel in einer einzelnen Tabelle

[Sehen Sie sich an, wie Sudhi Ramamurthy dieses Beispiel auf YouTube durchspazieren.](https://youtu.be/di-8JukK3Lc)
