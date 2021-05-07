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
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="b44cc-103">Kombinieren von Daten aus Excel Tabellen in einer einzelnen Tabelle</span><span class="sxs-lookup"><span data-stu-id="b44cc-103">Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="b44cc-104">In diesem Beispiel werden Daten aus Excel Tabellen in einer einzigen Tabelle kombiniert, die alle Zeilen enthält.</span><span class="sxs-lookup"><span data-stu-id="b44cc-104">This sample combines data from multiple Excel tables into a single table that includes all the rows.</span></span> <span data-ttu-id="b44cc-105">Es wird davon ausgegangen, dass alle verwendeten Tabellen dieselbe Struktur haben.</span><span class="sxs-lookup"><span data-stu-id="b44cc-105">It assumes that all tables being used have the same structure.</span></span>

<span data-ttu-id="b44cc-106">Es gibt zwei Variationen dieses Skripts:</span><span class="sxs-lookup"><span data-stu-id="b44cc-106">There are two variations of this script:</span></span>

1. <span data-ttu-id="b44cc-107">Das [erste Skript](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) kombiniert alle Tabellen in der Excel Datei.</span><span class="sxs-lookup"><span data-stu-id="b44cc-107">The [first script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combines all tables in the Excel file.</span></span>
1. <span data-ttu-id="b44cc-108">Das [zweite Skript](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) ruft tabellen selektiv innerhalb einer Gruppe von Arbeitsblättern ab.</span><span class="sxs-lookup"><span data-stu-id="b44cc-108">The [second script](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) selectively gets tables within a set of worksheets.</span></span>

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="b44cc-109">Beispielcode: Kombinieren von Daten aus Excel Tabellen in einer einzelnen Tabelle</span><span class="sxs-lookup"><span data-stu-id="b44cc-109">Sample code: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="b44cc-110">Laden Sie die Beispieldatei <a href="tables-copy.xlsx">tables-copy.xlsx</a> und verwenden Sie sie mit dem folgenden Skript, um sie selbst auszuprobieren!</span><span class="sxs-lookup"><span data-stu-id="b44cc-110">Download the sample file <a href="tables-copy.xlsx">tables-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

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

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a><span data-ttu-id="b44cc-111">Beispielcode: Kombinieren von Daten aus mehreren Excel tabellen in ausgewählten Arbeitsblättern in einer einzelnen Tabelle</span><span class="sxs-lookup"><span data-stu-id="b44cc-111">Sample code: Combine data from multiple Excel tables in select worksheets into a single table</span></span>

<span data-ttu-id="b44cc-112">Laden Sie die Beispieldatei <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> und verwenden Sie sie mit dem folgenden Skript, um sie selbst auszuprobieren!</span><span class="sxs-lookup"><span data-stu-id="b44cc-112">Download the sample file <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

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

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="b44cc-113">Schulungsvideo: Kombinieren von Daten aus mehreren Excel in einer einzelnen Tabelle</span><span class="sxs-lookup"><span data-stu-id="b44cc-113">Training video: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="b44cc-114">[Sehen Sie sich an, wie Sudhi Ramamurthy dieses Beispiel auf YouTube durchspazieren.](https://youtu.be/di-8JukK3Lc)</span><span class="sxs-lookup"><span data-stu-id="b44cc-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/di-8JukK3Lc).</span></span>
